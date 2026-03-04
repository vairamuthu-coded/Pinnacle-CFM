using CrystalDecisions.Shared;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Pinnacle.Transactions.PIN
{
    public partial class DateTimeReport : Form,ToolStripAccess
    {
        private static DateTimeReport _instance;
        DataTable dt; DataSet ds;
        Models.CommonClass com = new Models.CommonClass();
        public static DateTimeReport Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new DateTimeReport();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }
        public DateTimeReport()
        {
            InitializeComponent();
            GlobalVariables.HeaderName.Text = Class.Users.ScreenName; GlobalVariables.CurrentForm = this;
        }
        DataTable dtprint = new DataTable(); DataTable dt2 = new DataTable();
        private void DateTimeReport_Load(object sender, EventArgs e)
        {
            string sel = "SELECT  DISTINCT IDCARDNO FROM ASPTBLEMP A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID  and a.active='T' WHERE B.COMPCODE='"+Class.Users.HCompcode+"' order by 1";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLEMP");
            comboidcardno.DataSource = ds.Tables[0];
            comboidcardno.ValueMember = "IDCARDNO";
            comboidcardno.DisplayMember = "IDCARDNO";
            var lastDay = DateTime.DaysInMonth(System.DateTime.Now.Year, System.DateTime.Now.Month);
            var first = new DateTime(System.DateTime.Now.Year, System.DateTime.Now.Month, 1);
            var first1 = new DateTime(System.DateTime.Now.Year, System.DateTime.Now.Month, lastDay);
            frmdate.Value = first;
            todate.Value = first1;
            News();
        }
        private DateTime MonthYearDT()
        {
            frmdate.Format = DateTimePickerFormat.Custom;
            frmdate.CustomFormat = "dd-MM-yyyy";
            frmdate.Enabled = true;
            return frmdate.Value;
        }
        private DateTime MonthYearDT1()
        {
            todate.Format = DateTimePickerFormat.Custom;
            todate.CustomFormat = "dd-MM-yyyy";
            todate.Enabled = true;
            return todate.Value;
        }
        private void comboidcardno_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (Convert.ToInt64(comboidcardno.Text) > 0)
                {
                    string sel = "SELECT  DISTINCT empname FROM ASPTBLEMP A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID WHERE B.COMPCODE='" + Class.Users.HCompcode + "' AND A.IDCARDNO='" + comboidcardno.Text + "' ";
                    DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLEMP");
                    DataTable dt = ds.Tables["ASPTBLEMP"];
                    txtempname.Text = dt.Rows[0]["empname"].ToString();

                }
            }
            catch (Exception ex) { }
        }
      Report.PIN.DateTimeReport rd = new Report.PIN.DateTimeReport();
       Report.PIN.Attendance rd1 = new Report.PIN.Attendance();
        Report.PIN.IDCardWise rd2 = new Report.PIN.IDCardWise();
        Report.PIN.DateWise rd3 = new Report.PIN.DateWise();
        private DataTable GetTransposedTable(DataTable dt)
        {
            DataTable newTable = new DataTable();
            newTable.Columns.Add(new DataColumn("0", typeof(string)));
            for (int i = 0; i < dt.Columns.Count; i++)
            {
                DataRow newRow = newTable.NewRow();
                newRow[0] = dt.Columns[i].ColumnName;
                for (int j = 1; j <= dt.Rows.Count; j++)
                {
                    if (newTable.Columns.Count < dt.Rows.Count + 1)
                        newTable.Columns.Add(new DataColumn(j.ToString(), typeof(string)));
                    newRow[j] = dt.Rows[j - 1][i];
                }
                newTable.Rows.Add(newRow);
            }
            return newTable;
        }

        TimeSpan totaltime; TimeSpan totalhrs; TimeSpan permissionhrs; TimeSpan remainhrs;
        TimeSpan totaltime1; string[] STR1;
        private void butView_Click(object sender, EventArgs e)
        {
            try
            {

                Class.Users.UserTime = 0; Int64 ho = 0, da = 0;
                Int64 mi = 0; TimeSpan latehours1, latehours2, enttime10; TimeSpan pertime;
                Int64 se = 0; string val1 = "";
                Int64 totmits = 0; Int64 totmits1 = 0, totmits2 = 0, totmits3 = 0;
                DataTable dtmonth = new DataTable(); string com = "", add = "", idcard = "", empname = "";
                dtmonth.Columns.Add("month", typeof(string));
                dtmonth.Columns.Add("officetime", typeof(string));
                dtmonth.Columns.Add("lunchtime", typeof(string));
                dtmonth.Columns.Add("teatime", typeof(string));
                dtmonth.Columns.Add("teatime1", typeof(string));
                string sm = frmdate.Value.ToString().Substring(3, 2);
                string sm1 = monthname(sm);
                dtmonth.Rows.Add(sm1 + "     " + "(" + frmdate.Value.ToString("dd-MM-yyyy") + "  TO " + todate.Value.ToString("dd-MM-yyyy") + ")");
                int sno = 0;
                if (checkBox1.Checked == true)
                {
                    dt2.Rows.Clear(); News();
                    dt2.Rows.Clear(); int cnt = 0; crystalReportViewer1.ReportSource = null; crystalReportViewer1.Refresh();//b.asptblempid,b.idcardno
                    string sel0 = "SELECT distinct '' as asptblempid,b.idcardno FROM " + Class.Users.HCompcode + "TRS_ATTLOG a join asptblemp b on a.enrollno=b.idcardno join gtcompmast c on c.gtcompmastid = b.compcode where b.active='T' AND  c.compcode='" + Class.Users.HCompcode + "'  order by 2";
                    DataSet ds0 = Utility.ExecuteSelectQuery(sel0, "asptblemp");
                    DataTable dt0 = ds0.Tables["asptblemp"];
                    if (dt0.Rows.Count > 0)
                    {
                        for (int a = 0; a < dt0.Rows.Count; a++)
                        {
                            if (dt0.Rows.Count > 0)
                            {
                                sno = 0;
                                permissionhrs = TimeSpan.Zero; remainhrs = TimeSpan.Zero; totalhrs = TimeSpan.Zero; totaltime = TimeSpan.Zero;
                                string sel1 = "SELECT distinct '' asptblempid,b.idcardno,b.empname,DATE_FORMAT(a.indate,'%Y-%m-%d') as indate,'' minimumin,'' mteaout,'' mteain,'' lunchout,'' lucnchin,'' eveteaout,'' eveteain,''per0,''per1,''per2,''per3,''per4,''per5,''per6,''per7,''per8,''per9,'' maxiumout,'' timediffer,'' mteaexcess,'' luncexcess,'' eventeaexcess,'' totexcess,c.compname,c.address,C.PERMISSIONHRS,'' totalhrs,'' outtime,c.intime as starttime,c.outtime as endtime,'' datetimerecord,''per10,'' as per11,'' as per12,b.gender,c.femaleout   FROM " + Class.Users.HCompcode + "TRS_ATTLOG a join asptblemp b on a.enrollno=b.idcardno join gtcompmast c on c.gtcompmastid = b.compcode where c.compcode='" + Class.Users.HCompcode + "'  and b.idcardno ='" + dt0.Rows[a]["idcardno"].ToString() + "' and a.indate between '" + frmdate.Value.ToString("yyyy-MM-dd") + "' and  '" + todate.Value.ToString("yyyy-MM-dd") + "' order by 4,2";
                                DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "ASPTBLEMP");
                                DataTable dt1 = ds1.Tables["ASPTBLEMP"];
                                if (dt1.Rows.Count > 0)
                                {
                                    TimeSpan offtime = TimeSpan.Parse(dt1.Rows[0]["starttime"].ToString());
                                    TimeSpan offentime = new TimeSpan();
                                    if (dt1.Rows[0]["gender"].ToString() == "F")
                                    {
                                        offentime = TimeSpan.Parse(dt1.Rows[0]["femaleout"].ToString());
                                    }
                                    else
                                    {
                                        offentime = TimeSpan.Parse(dt1.Rows[0]["endtime"].ToString());
                                    }

                                    for (int i = 0; i < dt1.Rows.Count; i++)
                                    {
                                        sno += 1;
                                        string sel3 = "SELECT distinct a.enrollno as idcardno, min(a.intime) as intime,max(a.outtime) as outtime FROM " + Class.Users.HCompcode + "TRS_ATTLOG a join asptblemp b on a.enrollno=b.idcardno join gtcompmast c on c.gtcompmastid = b.compcode where c.compcode='" + Class.Users.HCompcode + "' and b.idcardno ='" + dt1.Rows[i]["idcardno"].ToString() + "' and a.indate='" + Convert.ToDateTime(dt1.Rows[i]["indate"].ToString()).ToString("yyyy-MM-dd") + "' group by a.enrollno order by 1 ";
                                        DataSet ds3 = Utility.ExecuteSelectQuery(sel3, "ASPTBLEMP");
                                        DataTable dt3 = ds3.Tables["ASPTBLEMP"];

                                        string sell = "select a.asptblmanualID,a.manualdate,a.permission, b.empname , b.idcardno  , a.active    from  asptblmanual a  join asptblemp b on a.empname=b.asptblempid join gtcompmast c on c.gtcompmastid=b.compcode   where c.compcode='" + Class.Users.HCompcode + "' and  b.idcardno ='" + dt1.Rows[i]["idcardno"].ToString() + "' and a.manualdate='" + dt1.Rows[i]["indate"].ToString() + "' order by 1";
                                        DataSet dsl = Utility.ExecuteSelectQuery(sell, "asptblmanual");
                                        DataTable dtl = dsl.Tables["asptblmanual"];

                                        string sel7 = "SELECT max(a.outtime) as intime,c.PERMISSIONHRS  FROM " + Class.Users.HCompcode + "TRS_ATTLOG a join asptblemp b on a.enrollno=b.idcardno join gtcompmast c on c.gtcompmastid = b.compcode where c.compcode='" + Class.Users.HCompcode + "' and b.idcardno ='" + dt1.Rows[i]["idcardno"].ToString() + "' and a.indate='" + Convert.ToDateTime(dt1.Rows[i]["indate"].ToString()).ToString("yyyy-MM-dd") + "' group by c.PERMISSIONHRS order by 1 ";
                                        DataSet ds7 = Utility.ExecuteSelectQuery(sel7, "ASPTBLEMP");
                                        DataTable dt7 = ds7.Tables["ASPTBLEMP"];

                                        STR1 = null;
                                        if (dtl.Rows.Count > 0)
                                        {
                                            STR1 = dtl.Rows[0]["permission"].ToString().Split(':');
                                        }
                                        if (dt2.Rows.Count <= 0)
                                        {
                                            dt2 = dt1.Clone();
                                            cnt = 0;


                                        }
                                        dt2.Rows.Add(); totaltime1 = TimeSpan.Zero;
                                        int totalcount = dt2.Rows.Count; int totalcount3 = dt3.Rows.Count;
                                        for (int j = 0; j < dt3.Rows.Count; j++)
                                        {

                                            if (dt2.Rows.Count == totalcount)
                                            {
                                                if (j == 0)
                                                {

                                                    TimeSpan starttime = TimeSpan.Parse(dt1.Rows[i]["starttime"].ToString());
                                                    TimeSpan starttime1 = TimeSpan.Parse(dt3.Rows[j]["intime"].ToString());
                                                    if (offtime > starttime1)
                                                    {
                                                        dt2.Rows[cnt + i]["minimumin"] = starttime1;
                                                      

                                                    }
                                                    else
                                                    {
                                                        dt2.Rows[cnt + i]["minimumin"] = dt3.Rows[j]["intime"].ToString();
                                                        TimeSpan startend = TimeSpan.Parse(dt1.Rows[i]["starttime"].ToString());
                                                        dt2.Rows[cnt + i]["starttime"] = dt3.Rows[j]["intime"].ToString();                                                   
                                                        TimeSpan endtime = TimeSpan.Parse(dt3.Rows[j]["outtime"].ToString());
                                                        TimeSpan differ = endtime.Subtract(startend);
                                                        TimeSpan differ1 = endtime - offtime;
                                                        dt2.Rows[cnt + i]["per0"] = differ1.ToString();
                                                        totaltime += differ1;
                                                        totalhrs += differ1;
                                                        totaltime1 += differ1;
                                                        dt2.Rows[cnt + i]["totexcess"] = totaltime.ToString();
                                                    }

                                                    dt2.Rows[cnt + i]["asptblempid"] = sno;
                                                    dt2.Rows[cnt + i]["indate"] = Convert.ToDateTime(dt1.Rows[i]["indate"].ToString()).ToString("dd-MM-yyyy");
                                                    dt2.Rows[cnt + i]["compname"] = dt1.Rows[i]["compname"].ToString();
                                                    dt2.Rows[cnt + i]["address"] = dt1.Rows[i]["address"].ToString();
                                                    dt2.Rows[cnt + i]["idcardno"] = dt1.Rows[i]["idcardno"].ToString();
                                                    dt2.Rows[cnt + i]["empname"] = dt1.Rows[i]["empname"].ToString();
                                                    permissionhrs += TimeSpan.Parse(dt1.Rows[cnt + i]["PERMISSIONHRS"].ToString());
                                                    dt2.Rows[cnt + i]["starttime"] = dt1.Rows[i]["starttime"].ToString();
                                                    if (dt1.Rows[i]["gender"].ToString() == "F")
                                                    {
                                                        dt2.Rows[cnt + i]["endtime"] = dt1.Rows[i]["femaleout"].ToString();
                                                    }
                                                    else
                                                    {
                                                        dt2.Rows[cnt + i]["endtime"] = dt1.Rows[i]["endtime"].ToString();
                                                    }
                                                    com = dt1.Rows[i]["compname"].ToString();
                                                    add = dt1.Rows[i]["address"].ToString();
                                                    idcard = dt1.Rows[i]["idcardno"].ToString();
                                                    empname = dt1.Rows[i]["empname"].ToString();

                                                }

                                                if (j == 1)
                                                {
                                                    dt2.Rows[cnt + i]["mteaout"] = dt3.Rows[j]["intime"].ToString();
                                                    dt2.Rows[cnt]["idcardno"] = dt1.Rows[i]["idcardno"].ToString();
                                                    dt2.Rows[cnt]["empname"] = dt1.Rows[i]["empname"].ToString();
                                                }

                                                if (j == 2)
                                                {
                                                    dt2.Rows[cnt + i]["mteain"] = dt3.Rows[j]["intime"].ToString();
                                                    dt2.Rows[cnt + i]["idcardno"] = dt1.Rows[i]["idcardno"].ToString();
                                                    dt2.Rows[cnt + i]["empname"] = dt1.Rows[i]["empname"].ToString();


                                                    if (dt2.Rows[cnt + i]["mteain"].ToString() != "")
                                                    {
                                                        DateTime statetime = DateTime.Parse(dt2.Rows[cnt + i]["mteaout"].ToString());
                                                        DateTime endtime = DateTime.Parse(dt2.Rows[cnt + i]["mteain"].ToString());
                                                        TimeSpan differ = endtime.Subtract(statetime);
                                                        TimeSpan differ1 = endtime - statetime;
                                                        dt2.Rows[cnt + i]["mteaexcess"] = differ1.ToString();
                                                        totaltime = differ1;
                                                        totalhrs += differ1; totaltime1 += differ1;
                                                    }
                                                }

                                                if (j == 3)
                                                {
                                                    dt2.Rows[cnt + i]["lunchout"] = dt3.Rows[j]["intime"].ToString();
                                                    dt2.Rows[cnt + i]["idcardno"] = dt1.Rows[i]["idcardno"].ToString();
                                                    dt2.Rows[cnt + i]["empname"] = dt1.Rows[i]["empname"].ToString();

                                                }

                                                if (j == 4)
                                                {
                                                    dt2.Rows[cnt + i]["lucnchin"] = dt3.Rows[j]["intime"].ToString();
                                                    dt2.Rows[cnt + i]["idcardno"] = dt1.Rows[i]["idcardno"].ToString();
                                                    dt2.Rows[cnt + i]["empname"] = dt1.Rows[i]["empname"].ToString();
                                                    if (dt2.Rows[cnt + i]["lucnchin"].ToString() != "")
                                                    {
                                                        DateTime statetime = DateTime.Parse(dt2.Rows[cnt + i]["lunchout"].ToString());
                                                        DateTime endtime = DateTime.Parse(dt2.Rows[cnt + i]["lucnchin"].ToString());
                                                        TimeSpan differ = endtime.Subtract(statetime);
                                                        TimeSpan differ1 = endtime - statetime;
                                                        dt2.Rows[cnt + i]["luncexcess"] = differ1.ToString();
                                                        totaltime += differ1;
                                                        totalhrs += differ1; totaltime1 += differ1;
                                                    }

                                                }

                                                if (j == 5)
                                                {
                                                    dt2.Rows[cnt + i]["eveteaout"] = dt3.Rows[j]["intime"].ToString();
                                                    dt2.Rows[cnt + i]["idcardno"] = dt1.Rows[i]["idcardno"].ToString();
                                                    dt2.Rows[cnt + i]["empname"] = dt1.Rows[i]["empname"].ToString();

                                                }

                                                if (j == 6)
                                                {
                                                    dt2.Rows[cnt + i]["eveteain"] = dt3.Rows[j]["intime"].ToString();
                                                    dt2.Rows[cnt + i]["idcardno"] = dt1.Rows[i]["idcardno"].ToString();
                                                    dt2.Rows[cnt + i]["empname"] = dt1.Rows[i]["empname"].ToString();
                                                    if (dt2.Rows[cnt + i]["eveteain"].ToString() != "")
                                                    {
                                                        DateTime statetime = DateTime.Parse(dt2.Rows[cnt + i]["eveteaout"].ToString());
                                                        DateTime endtime = DateTime.Parse(dt2.Rows[cnt + i]["eveteain"].ToString());
                                                        TimeSpan differ = endtime.Subtract(statetime);
                                                        TimeSpan differ1 = endtime - statetime;
                                                        dt2.Rows[cnt + i]["eventeaexcess"] = differ1.ToString();
                                                        totaltime += differ1;
                                                        totalhrs += differ1; totaltime1 += differ1;
                                                        dt2.Rows[cnt + i]["totexcess"] = totaltime.ToString();
                                                    }
                                                }

                                                if (j == 7)
                                                {
                                                    dt2.Rows[cnt + i]["per1"] = dt3.Rows[j]["intime"].ToString();
                                                    dt2.Rows[cnt + i]["idcardno"] = dt1.Rows[i]["idcardno"].ToString();
                                                    dt2.Rows[cnt + i]["empname"] = dt1.Rows[i]["empname"].ToString();

                                                }

                                                if (j == 8)
                                                {
                                                    dt2.Rows[cnt + i]["per2"] = dt3.Rows[j]["intime"].ToString();
                                                    dt2.Rows[cnt + i]["idcardno"] = dt1.Rows[i]["idcardno"].ToString();
                                                    dt2.Rows[cnt + i]["empname"] = dt1.Rows[i]["empname"].ToString();
                                                    if (dt2.Rows[cnt + i]["per2"].ToString() != "")
                                                    {
                                                        DateTime statetime = DateTime.Parse(dt2.Rows[cnt + i]["per1"].ToString());
                                                        DateTime endtime = DateTime.Parse(dt2.Rows[cnt + i]["per2"].ToString());
                                                        TimeSpan differ = endtime.Subtract(statetime);
                                                        TimeSpan differ1 = endtime - statetime;
                                                        dt2.Rows[cnt + i]["per3"] = differ1.ToString();
                                                        totaltime += differ1;
                                                        totalhrs += differ1; totaltime1 += differ1;
                                                        dt2.Rows[cnt + i]["totexcess"] = totaltime.ToString();


                                                    }
                                                }

                                                if (j == 9)
                                                {
                                                    dt2.Rows[cnt + i]["per4"] = dt3.Rows[j]["intime"].ToString();
                                                    dt2.Rows[cnt + i]["idcardno"] = dt1.Rows[i]["idcardno"].ToString();
                                                    dt2.Rows[cnt + i]["empname"] = dt1.Rows[i]["empname"].ToString();

                                                }

                                                if (j == 10)
                                                {
                                                    dt2.Rows[cnt + i]["per5"] = dt3.Rows[j]["intime"].ToString();
                                                    dt2.Rows[cnt + i]["idcardno"] = dt1.Rows[i]["idcardno"].ToString();
                                                    dt2.Rows[cnt + i]["empname"] = dt1.Rows[i]["empname"].ToString();
                                                    if (dt2.Rows[cnt + i]["eveteain"].ToString() != "")
                                                    {
                                                        DateTime statetime = DateTime.Parse(dt2.Rows[cnt + i]["per4"].ToString());
                                                        DateTime endtime = DateTime.Parse(dt2.Rows[cnt + i]["per5"].ToString());
                                                        TimeSpan differ = endtime.Subtract(statetime);
                                                        TimeSpan differ1 = endtime - statetime;
                                                        dt2.Rows[cnt + i]["per6"] = differ1.ToString();
                                                        totaltime += differ1;
                                                        totalhrs += differ1; totaltime1 += differ1;
                                                        dt2.Rows[cnt + i]["totexcess"] = totaltime.ToString();


                                                    }
                                                }
                                                if (j == 11)
                                                {
                                                    dt2.Rows[cnt + i]["per7"] = dt3.Rows[j]["intime"].ToString();
                                                    dt2.Rows[cnt + i]["idcardno"] = dt1.Rows[i]["idcardno"].ToString();
                                                    dt2.Rows[cnt + i]["empname"] = dt1.Rows[i]["empname"].ToString();

                                                }

                                                if (j == 12)
                                                {
                                                    dt2.Rows[cnt + i]["per8"] = dt3.Rows[j]["intime"].ToString();
                                                    dt2.Rows[cnt + i]["idcardno"] = dt1.Rows[i]["idcardno"].ToString();
                                                    dt2.Rows[cnt + i]["empname"] = dt1.Rows[i]["empname"].ToString();
                                                    if (dt2.Rows[cnt + i]["eveteain"].ToString() != "")
                                                    {
                                                        DateTime statetime = DateTime.Parse(dt2.Rows[cnt + i]["per7"].ToString());
                                                        DateTime endtime = DateTime.Parse(dt2.Rows[cnt + i]["per8"].ToString());
                                                        TimeSpan differ = endtime.Subtract(statetime);
                                                        TimeSpan differ1 = endtime - statetime;
                                                        dt2.Rows[cnt + i]["per9"] = differ1.ToString();
                                                        totaltime += differ1;
                                                        totalhrs += differ1; totaltime1 += differ1;
                                                        dt2.Rows[cnt + i]["totexcess"] = totaltime.ToString();

                                                    }
                                                }
                                                if (j == 0 && totalcount3 == 1 || j == 1 && totalcount3 == 1 || j == 2 && totalcount3 == 1 || j == 3 && totalcount3 == 1 || j == 4 && totalcount3 == 1 || j == 5 && totalcount3 == 1 || j == 6 && totalcount3 == 1 || j == 7 && totalcount3 == 1 || j == 8 && totalcount3 == 1 || j == 9 && totalcount3 == 1 || j == 10 && totalcount3 == 1 || j == 11 && totalcount3 == 1 || j == 12 && totalcount3 == 1 || j == 13)
                                                {
                                                    //  MessageBox.Show(dt1.Rows[i]["indate"].ToString() + "  "+ dt1.Rows[i]["idcardno"].ToString() + " " + dt7.Rows[0]["intime"].ToString());
                                                    TimeSpan offstarttime1; TimeSpan offstarttime2;
                                                    TimeSpan diff;
                                                    dt2.Rows[cnt + i]["idcardno"] = dt1.Rows[i]["idcardno"].ToString();
                                                    dt2.Rows[cnt + i]["empname"] = dt1.Rows[i]["empname"].ToString();
                                                    if (dt2.Rows[cnt + i]["minimumin"].ToString() != "")
                                                    {
                                                        TimeSpan minimumin = TimeSpan.Parse(dt2.Rows[cnt + i]["minimumin"].ToString());
                                                        TimeSpan mintime = TimeSpan.Parse(dt2.Rows[cnt+i]["starttime"].ToString());//minimumin dt2.Rows[cnt + i]["starttime"].ToString()
                                                        TimeSpan maxtime = TimeSpan.Parse(dt7.Rows[0]["intime"].ToString());
                                                        TimeSpan differ = maxtime.Subtract(minimumin);
                                                        TimeSpan endtime1 = TimeSpan.Parse(dt2.Rows[cnt + i]["endtime"].ToString());

                                                        //offstarttime1 = offtime.Subtract(mintime);
                                                       //  offstarttime2 = maxtime.Subtract(offentime);
                                                       
                                                        //diff = minimumin.Subtract(mintime);
                                                        //differ = maxtime.Subtract(mintime);

                                                        // if (diff.Hours > 0 || diff.Minutes > 0 || diff.Seconds > 0)
                                                        // {
                                                        if (maxtime.Hours >= 19)
                                                        {
                                                            if (minimumin.Hours >= 09 && maxtime.Hours >= 19) {
                                                                differ = endtime1.Subtract(offtime);
                                                                offstarttime1 = offtime.Subtract(minimumin);
                                                                offstarttime2 = offentime.Subtract(endtime1);
                                                            } else {
                                                                //differ = endtime1.Subtract(minimumin);
                                                                //offstarttime1 = offtime.Subtract(minimumin);
                                                                //offstarttime2 = offentime.Subtract(endtime1);
                                                                differ = maxtime.Subtract(minimumin);
                                                                offstarttime1 = offtime.Subtract(offtime);
                                                                offstarttime2 = offentime.Subtract(endtime1);
                                                            }
                                                            
                                                        }
                                                        else
                                                        {
                                                            //differ = maxtime.Subtract(mintime);
                                                            differ = endtime1.Subtract(minimumin);
                                                            offstarttime1 = offtime.Subtract(minimumin);
                                                            offstarttime2 = endtime1.Subtract(maxtime);
                                                        }
                                                        val1 = offstarttime1.ToString().Substring(0, 1);                                                  

                                                        if (combotype.Text == "Attendance")
                                                        {
                                                            da = 0; ho = 0; mi = 0; se = 0; totmits = 0; totmits1 = 0; totmits2 = 0;
                                                            da = offstarttime1.Days;
                                                            ho = offstarttime1.Hours * 60;
                                                            mi = offstarttime1.Minutes;
                                                            totmits = da + ho + mi;
                                                            totmits1 = totmits;
                                                            da = 0; ho = 0; mi = 0; se = 0; totmits = 0;
                                                            da = offstarttime2.Days;
                                                            ho = offstarttime2.Hours * 60;
                                                            mi = offstarttime2.Minutes;
                                                            totmits = da + ho + mi;
                                                            totmits2 = totmits;
                                                            if (offstarttime1.Minutes <= 0)
                                                            {

                                                                val1 = totmits1.ToString().Trim('-');
                                                                TimeSpan result1 = TimeSpan.FromMinutes(Convert.ToInt64(val1));
                                                                da = 0; ho = 0; mi = 0; se = 0; totmits = 0; totmits1 = 0; totmits1 = 0;
                                                                da = result1.Days;
                                                                ho = result1.Hours * 60;
                                                                mi = result1.Minutes;
                                                                totmits = da + ho + mi; totmits1 = 0;
                                                                totmits1 = totmits;
                                                                da = 0; ho = 0; mi = 0; se = 0; totmits = 0; val1 = "";
                                                                if (offstarttime2.Ticks > 0)
                                                                {
                                                                    //totmits2 = 0; totmits = 0;
                                                                }
                                                                else
                                                                {
                                                                    val1 = totmits2.ToString().Trim('-');
                                                                    TimeSpan result2 = TimeSpan.FromMinutes(Convert.ToInt64(val1));
                                                                    da = result2.Days;
                                                                    ho = result2.Hours * 60;
                                                                    mi = result2.Minutes;
                                                                    totmits = da + ho + mi; totmits2 = 0;
                                                                    totmits2 = totmits;
                                                                }
                                                                val1 = "";


                                                            }
                                                            Int64 tiemdiff = totmits2 + totmits1;
                                                            val1 = tiemdiff.ToString().Trim('-');
                                                            TimeSpan result = TimeSpan.FromMinutes(Convert.ToInt64(val1));
                                                            string fromTimeString = result.ToString("hh':'mm");
                                                            ho = 0; mi = 0;
                                                            if (dtl.Rows.Count > 0)
                                                            {
                                                                if (STR1.Length > 1)
                                                                {
                                                                    string Z1 = STR1[0];
                                                                    string Z2 = STR1[1];
                                                                    if (Convert.ToInt64(Z1.ToString()) > 0 || Convert.ToInt64(Z2.ToString()) > 0)
                                                                    {
                                                                        pertime = TimeSpan.Parse(dtl.Rows[0]["permission"].ToString());
                                                                        ho = pertime.Hours * 60;
                                                                        mi = pertime.Minutes;
                                                                        Int64 tot = ho + mi;
                                                                        tiemdiff += tot; pertime = TimeSpan.Zero;
                                                                    }
                                                                }
                                                            }
                                                            dt2.Rows[cnt + i]["per1"] = tiemdiff.ToString();
                                                            dt2.Rows[cnt + i]["per10"] = fromTimeString.ToString();
                                                        }
                                                        if (dt7.Rows[0]["intime"].ToString() != "")
                                                        {

                                                            TimeSpan enttime = TimeSpan.Parse(dt1.Rows[i]["endtime"].ToString());
                                                            TimeSpan enttime1 = TimeSpan.Parse(dt7.Rows[0]["intime"].ToString());
                                                            if (offentime > enttime1)
                                                            {
                                                                dt2.Rows[cnt + i]["maxiumout"] = dt7.Rows[0]["intime"].ToString();
                                                                DateTime statetime = DateTime.Parse(dt2.Rows[cnt + i]["minimumin"].ToString());
                                                                DateTime endtime=new DateTime(); TimeSpan differ2 = new TimeSpan(); TimeSpan differ1 =new TimeSpan();
                                                                if (dt1.Rows[0]["gender"].ToString() == "F")
                                                                {
                                                                    endtime = DateTime.Parse(dt1.Rows[0]["femaleout"].ToString());
                                                                }
                                                                else
                                                                {
                                                                    if (dt3.Rows[j]["outtime"].ToString() == "0") { }
                                                                    else
                                                                    {
                                                                        DateTime starttime1 = DateTime.Parse(dt2.Rows[cnt + i]["minimumin"].ToString());
                                                                        DateTime starttime2 = DateTime.Parse(dt2.Rows[cnt + i]["starttime"].ToString());
                                                                        if (starttime1.Hour < 9) {
                                                                           
                                                                            endtime = DateTime.Parse(dt3.Rows[j]["outtime"].ToString());
                                                                            differ2 = endtime.Subtract(starttime2);
                                                                            differ1 = endtime - starttime2;
                                                                            dt2.Rows[cnt + i]["timediffer"] = differ1;
                                                                        }
                                                                        else
                                                                        {
                                                                            endtime = DateTime.Parse(dt3.Rows[j]["outtime"].ToString());
                                                                            differ2 = endtime.Subtract(starttime1);
                                                                            differ1 = endtime - starttime1;
                                                                            dt2.Rows[cnt + i]["timediffer"] = differ1;
                                                                        }
                                                                    }
                                                                   
                                                                }                                                                
                                                                dt2.Rows[cnt + i]["totexcess"] = totaltime.ToString();
                                                                if (dtl.Rows.Count > 0)
                                                                {
                                                                    if (STR1.Length > 1)
                                                                    {
                                                                        string Z1 = STR1[0];
                                                                        string Z2 = STR1[1];
                                                                        if (Convert.ToInt64(Z1.ToString()) > 0 || Convert.ToInt64(Z2.ToString()) > 0)
                                                                        {
                                                                            dt2.Rows[i]["per11"] = dtl.Rows[0]["permission"].ToString(); dtl.Rows.Clear();

                                                                        }
                                                                    }
                                                                }
                                                                if (dt2.Rows[cnt + i]["per11"].ToString() != "")
                                                                {
                                                                    TimeSpan enttime2 = TimeSpan.Parse(dt2.Rows[cnt + i]["per11"].ToString());
                                                                    enttime10 = TimeSpan.Parse(dt7.Rows[0]["PERMISSIONHRS"].ToString());
                                                                    dt2.Rows[cnt + i]["timediffer"] = differ1 + enttime2;
                                                                    dtl.Rows.Clear();
                                                                    TimeSpan timediff = TimeSpan.Parse(dt2.Rows[cnt + i]["timediffer"].ToString());
                                                                    dt2.Rows[cnt + i]["timediffer"] = timediff.ToString().Substring(0, 5);
                                                                    da = 0; ho = 0; mi = 0; totmits = 0;
                                                                    da = timediff.Days;
                                                                    ho = timediff.Hours * 60;
                                                                    mi = timediff.Minutes;
                                                                    totmits = da + ho + mi;
                                                                    dt2.Rows[cnt + i]["per12"] = totmits;

                                                                }
                                                                else
                                                                {
                                                                    dt2.Rows[cnt + i]["timediffer"] = differ1;
                                                                    enttime10 = TimeSpan.Parse(dt7.Rows[0]["PERMISSIONHRS"].ToString());
                                                                    TimeSpan timediff = TimeSpan.Parse(dt2.Rows[cnt + i]["timediffer"].ToString());
                                                                    dt2.Rows[cnt + i]["timediffer"] = timediff.ToString().Substring(0, 5);
                                                                    if (combotype.Text == "Attendance")
                                                                    {
                                                                        da = 0; ho = 0; mi = 0; totmits = 0;
                                                                        da = timediff.Days;
                                                                        ho = timediff.Hours * 60;
                                                                        mi = timediff.Minutes;
                                                                        totmits = da + ho + mi;
                                                                        dt2.Rows[cnt + i]["per12"] = totmits;
                                                                    }
                                                                    else
                                                                    {
                                                                        dt2.Rows[cnt + i]["per12"] = timediff;
                                                                    }
                                                                }
                                                            }
                                                            else
                                                            {
                                                                DateTime endtime, statetime, endtime2,statetime1; TimeSpan differ2 = new TimeSpan(); TimeSpan differ1 = new TimeSpan();
                                                                differ2 = TimeSpan.Zero; differ1 = TimeSpan.Zero;
                                                                if (dt3.Rows[j]["intime"].ToString() != "")
                                                                {
                                                                    dt2.Rows[cnt + i]["maxiumout"] = dt7.Rows[0]["intime"].ToString();// dt3.Rows[j]["outtime"].ToString();
                                                                    statetime = DateTime.Parse(dt2.Rows[cnt + i]["minimumin"].ToString());//minimumin
                                                                    endtime = DateTime.Parse(dt2.Rows[cnt + i]["maxiumout"].ToString());
                                                                   statetime1 = DateTime.Parse(dt2.Rows[cnt + i]["starttime"].ToString());
                                                                    if (endtime.Hour >= 19)
                                                                    {
                                                                        endtime2 = DateTime.Parse(dt2.Rows[cnt + i]["endtime"].ToString());                                                                        
                                                                        differ2 = endtime2.Subtract(statetime1);
                                                                        differ1 = endtime2 - statetime1;
                                                                    }
                                                                    else
                                                                    {
                                                                        differ2 = endtime.Subtract(statetime);
                                                                        differ1 = endtime - statetime;
                                                                    }
                                                                }

                                                                dt2.Rows[cnt + i]["timediffer"] = differ1;
                                                                if (dtl.Rows.Count > 0)
                                                                {
                                                                    dt2.Rows[i]["per11"] = dtl.Rows[0]["permission"].ToString(); dtl.Rows.Clear();
                                                                }
                                                                if (dt2.Rows[cnt + i]["per11"].ToString() != "")
                                                                {
                                                                    TimeSpan enttime2 = TimeSpan.Parse(dt2.Rows[cnt + i]["per11"].ToString());
                                                                    enttime10 = TimeSpan.Parse(dt7.Rows[0]["PERMISSIONHRS"].ToString());
                                                                    dt2.Rows[cnt + i]["timediffer"] = differ1 + enttime2;
                                                                    dtl.Rows.Clear();
                                                                    TimeSpan timediff = TimeSpan.Parse(dt2.Rows[cnt + i]["timediffer"].ToString());
                                                                    dt2.Rows[cnt + i]["timediffer"] = timediff.ToString().Substring(0, 5);
                                                                    if (combotype.Text == "Attendance")
                                                                    {
                                                                        da = 0; ho = 0; mi = 0; totmits = 0;
                                                                        da = timediff.Days;
                                                                        ho = timediff.Hours * 60;
                                                                        mi = timediff.Minutes;
                                                                        totmits = da + ho + mi;
                                                                        dt2.Rows[cnt + i]["per12"] = totmits;

                                                                    }
                                                                    else
                                                                    {
                                                                        dt2.Rows[cnt + i]["per12"] = timediff;
                                                                    }

                                                                }
                                                                else
                                                                {
                                                                    dt2.Rows[cnt + i]["timediffer"] = differ1;
                                                                    enttime10 = TimeSpan.Parse(dt7.Rows[0]["PERMISSIONHRS"].ToString());
                                                                    TimeSpan timediff = TimeSpan.Parse(dt2.Rows[cnt + i]["timediffer"].ToString());
                                                                    dt2.Rows[cnt + i]["timediffer"] = timediff.ToString().Substring(0, 5);
                                                                    if (combotype.Text == "Attendance")
                                                                    {
                                                                        da = 0; ho = 0; mi = 0; totmits = 0;
                                                                        da = timediff.Days;
                                                                        ho = timediff.Hours * 60;
                                                                        mi = timediff.Minutes;
                                                                        totmits = da + ho + mi;
                                                                        dt2.Rows[cnt + i]["per12"] = totmits;
                                                                    }
                                                                    else
                                                                    {
                                                                        dt2.Rows[cnt + i]["per12"] = timediff;
                                                                    }
                                                                }
                                                                if (dt2.Rows[cnt + i]["mteain"].ToString() == "")
                                                                {


                                                                }
                                                                else
                                                                {

                                                                    dt2.Rows[cnt + i]["totexcess"] = totaltime.ToString();

                                                                }
                                                            }

                                                        }
                                                    }
                                                }
                                                dt2.Rows[i]["totalhrs"] = totalhrs.ToString();
                                                dt2.Rows[i]["permissionhrs"] = permissionhrs.ToString();
                                                totalcount3--;
                                            }
                                        }
                                        TimeSpan rtot = TimeSpan.Parse(dt2.Rows[i]["totalhrs"].ToString());
                                        TimeSpan rper = TimeSpan.Parse(dt2.Rows[i]["permissionhrs"].ToString());
                                        TimeSpan rdiffer1 = rtot - rper;
                                        if (totaltime1.Ticks > 0)
                                        {
                                            dt2.Rows[i]["datetimerecord"] = totaltime1.ToString();
                                        }
                                        dt2.Rows[i]["outtime"] = rdiffer1.ToString();
                                    }
                                }


                            }
                            if (dt2.Rows.Count >= 1)
                            {
                                if (dtprint.Rows.Count >= 1)
                                {
                                    foreach (DataRow dr in dt2.Rows)
                                    {
                                        dtprint.Rows.Add(dr.ItemArray);
                                        dtprint.AcceptChanges();
                                    }
                                    dt2.Rows.Clear();
                                }
                                else
                                {
                                    var orderedRows = from row in dt2.AsEnumerable() select row;
                                    dtprint = orderedRows.CopyToDataTable(); dt2.Rows.Clear();

                                }
                            }

                        }

                    }
                    else
                    {
                        Cursor = Cursors.Default;
                        MessageBox.Show("No Data Found"); return;
                    }
                }
                else
                {
                    News();
                    sno = 0;
                    com = ""; add = ""; idcard = ""; empname = "";
                    dt2.Rows.Clear(); int cnt = 0; crystalReportViewer1.ReportSource = null; crystalReportViewer1.Refresh();
                    string sel10 = "select indate from (select adddate('1970-01-01',t4.i*10000 + t3.i*1000 + t2.i*100 + t1.i*10 + t0.i) indate from" +
 "(select 0 i union select 1 union select 2 union select 3 union select 4 union select 5 union select 6 union select 7 union select 8 union select 9) t0," +
 "(select 0 i union select 1 union select 2 union select 3 union select 4 union select 5 union select 6 union select 7 union select 8 union select 9) t1," +
 "(select 0 i union select 1 union select 2 union select 3 union select 4 union select 5 union select 6 union select 7 union select 8 union select 9) t2," +
 "(select 0 i union select 1 union select 2 union select 3 union select 4 union select 5 union select 6 union select 7 union select 8 union select 9) t3," +
 "(select 0 i union select 1 union select 2 union select 3 union select 4 union select 5 union select 6 union select 7 union select 8 union select 9) t4) ss " +
 " where indate between '" + frmdate.Value.ToString("yyyy-MM-dd") + "' and '" + todate.Value.ToString("yyyy-MM-dd") + "' order by 1 ";
                    DataSet ds10 = Utility.ExecuteSelectQuery(sel10, "dual");
                    DataTable dt0 = ds10.Tables["dual"];
                    if (dt0.Rows.Count > 0)
                    {

                        for (int m = 0; m < dt0.Rows.Count; m++)
                        {

                            var list = dt0.AsEnumerable();
                            string mindate = list.FirstOrDefault()["indate"].ToString();
                            string maxdate = list.LastOrDefault()["indate"].ToString();
                            string sel1 = "SELECT distinct '' asptblempid,DATE_FORMAT(a.indate,'%Y-%m-%d') as indate,'' minimumin,'' mteaout,'' mteain,'' lunchout,'' lucnchin,'' eveteaout,'' eveteain,''per0,''per1,''per2,''per3,''per4,''per5,''per6,''per7,''per8,''per9,'' maxiumout,'' timediffer,'' mteaexcess,'' luncexcess,'' eventeaexcess,'' totexcess,c.compname,c.address,b.idcardno,b.empname,C.PERMISSIONHRS,'' totalhrs,'' outtime,c.intime as starttime,c.outtime as endtime,'' datetimerecord,'' as per10,'' as per11,'' as per12,b.gender,c.femaleout  FROM " + Class.Users.HCompcode + "TRS_ATTLOG a join asptblemp b on a.enrollno=b.idcardno join gtcompmast c on c.gtcompmastid = b.compcode  where c.compcode='" + Class.Users.HCompcode + "' and  b.idcardno ='" + comboidcardno.Text + "' and a.indate between '" + dt0.Rows[m]["indate"].ToString() + "' and  '" + dt0.Rows[m]["indate"].ToString() + "' order by 2";
                            DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "ASPTBLEMP");
                            DataTable dt1 = ds1.Tables["ASPTBLEMP"];

                            if (dt1.Rows.Count > 0)
                            {
                                TimeSpan offtime = TimeSpan.Parse(dt1.Rows[0]["starttime"].ToString());
                                TimeSpan offentime = new TimeSpan();
                                if (dt1.Rows[0]["gender"].ToString() == "F")
                                {
                                    offentime = TimeSpan.Parse(dt1.Rows[0]["femaleout"].ToString());
                                }
                                else
                                {
                                    offentime = TimeSpan.Parse(dt1.Rows[0]["endtime"].ToString());
                                }

                                for (int i = 0; i < dt1.Rows.Count; i++)
                                {
                                    cnt = 0; sno += 1;
                                    string sel3 = "SELECT distinct  a.intime,a.outtime  FROM " + Class.Users.HCompcode + "TRS_ATTLOG a join asptblemp b on a.enrollno=b.idcardno join gtcompmast c on c.gtcompmastid = b.compcode where c.compcode='" + Class.Users.HCompcode + "' and  b.idcardno ='" + comboidcardno.Text + "' and a.indate='" + dt1.Rows[i]["indate"].ToString() + "'  order by 1 ";
                                    DataSet ds3 = Utility.ExecuteSelectQuery(sel3, "ASPTBLEMP");
                                    DataTable dt3 = ds3.Tables["ASPTBLEMP"];
                                    string sell = "select a.asptblmanualID,a.manualdate,a.permission, b.empname , b.idcardno  , a.active    from  asptblmanual a  join asptblemp b on a.empname=b.asptblempid  join gtcompmast c on c.gtcompmastid=b.compcode  where c.compcode='" + Class.Users.HCompcode + "' and  b.idcardno ='" + comboidcardno.Text + "' and a.manualdate='" + dt1.Rows[i]["indate"].ToString() + "' order by 1";
                                    DataSet dsl = Utility.ExecuteSelectQuery(sell, "asptblmanual");
                                    DataTable dtl = dsl.Tables["asptblmanual"];
                                    if (dt2.Rows.Count <= 0)
                                    {
                                        dt2 = dt1.Clone();
                                        cnt = 0; com = ""; add = ""; idcard = ""; empname = "";
                                    }
                                    dt2.Rows.Add();
                                    int totalcount = dt2.Rows.Count; int totalcount3 = dt3.Rows.Count;
                                    for (int j = 0; j < dt3.Rows.Count; j++)
                                    {
                                        if (dt2.Rows.Count == totalcount)
                                        {

                                            if (j == 0)
                                            {
                                                TimeSpan starttime1 = TimeSpan.Parse(dt3.Rows[j]["intime"].ToString());
                                                if (offtime > starttime1 || offtime == starttime1)
                                                {
                                                    // dt2.Rows[cnt + m]["starttime"] = offtime;
                                                    dt2.Rows[cnt + m]["minimumin"] = dt3.Rows[j]["intime"].ToString();// offtime;
                                                }
                                                else
                                                {
                                                    dt2.Rows[cnt + m]["minimumin"] = dt3.Rows[j]["intime"].ToString();
                                                    TimeSpan startend = TimeSpan.Parse(dt1.Rows[i]["starttime"].ToString());
                                                    dt2.Rows[cnt + m]["starttime"] = dt3.Rows[j]["intime"].ToString();//minimumin
                                                    TimeSpan endtime = TimeSpan.Parse(dt3.Rows[j]["intime"].ToString()); //TimeSpan.Parse(dt7.Rows[0]["intime"].ToString());
                                                    TimeSpan differ = endtime.Subtract(startend);
                                                    TimeSpan differ1 = endtime - offtime;
                                                    dt2.Rows[cnt + m]["per0"] = differ1.ToString();
                                                    totaltime += differ1;
                                                    totalhrs += differ1;
                                                    totaltime1 += differ1;
                                                    dt2.Rows[cnt + m]["totexcess"] = totaltime.ToString();
                                                    dt2.Rows[cnt + i]["totexcess"] = totaltime.ToString();
                                                    if (combotype.Text != "Attendance")
                                                    {
                                                        TimeSpan differ2 = starttime1 - offtime;
                                                        dt2.Rows[cnt + m]["per10"] = differ2.ToString().Substring(0, 5);

                                                    }
                                                }

                                                dt2.Rows[cnt + m]["asptblempid"] = sno;
                                                dt2.Rows[cnt + m]["indate"] = Convert.ToDateTime(dt1.Rows[i]["indate"].ToString()).ToString("dd-MM-yyyy");
                                                dt2.Rows[cnt]["compname"] = dt1.Rows[i]["compname"].ToString();
                                                com = dt1.Rows[i]["compname"].ToString();
                                                add = dt1.Rows[i]["address"].ToString();
                                                idcard = dt1.Rows[i]["idcardno"].ToString();
                                                empname = dt1.Rows[i]["empname"].ToString();
                                                dt2.Rows[cnt]["address"] = dt1.Rows[i]["address"].ToString();
                                                dt2.Rows[cnt + m]["idcardno"] = dt1.Rows[i]["idcardno"].ToString();
                                                dt2.Rows[cnt + m]["empname"] = dt1.Rows[i]["empname"].ToString();
                                                permissionhrs += TimeSpan.Parse(dt1.Rows[cnt + i]["PERMISSIONHRS"].ToString());
                                                dt2.Rows[cnt + m]["starttime"] = dt1.Rows[i]["starttime"].ToString();
                                                dt2.Rows[cnt + m]["endtime"] = dt1.Rows[i]["endtime"].ToString();
                                            }

                                            if (j == 1)
                                            {

                                                dt2.Rows[cnt + m]["mteaout"] = dt3.Rows[j]["intime"].ToString();
                                                dt2.Rows[cnt]["idcardno"] = dt1.Rows[i]["idcardno"].ToString();
                                                dt2.Rows[cnt]["empname"] = dt1.Rows[i]["empname"].ToString();
                                            }

                                            if (j == 2)
                                            {
                                                dt2.Rows[cnt + m]["mteain"] = dt3.Rows[j]["intime"].ToString();
                                                dt2.Rows[cnt + m]["idcardno"] = dt1.Rows[i]["idcardno"].ToString();
                                                dt2.Rows[cnt + m]["empname"] = dt1.Rows[i]["empname"].ToString();
                                                if (dt2.Rows[cnt + m]["mteain"].ToString() != "")
                                                {
                                                    DateTime statetime = DateTime.Parse(dt2.Rows[cnt + m]["mteaout"].ToString());
                                                    DateTime endtime = DateTime.Parse(dt2.Rows[cnt + m]["mteain"].ToString());
                                                    TimeSpan differ = endtime.Subtract(statetime);
                                                    TimeSpan differ1 = endtime - statetime;
                                                    dt2.Rows[cnt + i]["mteaexcess"] = differ1.ToString();
                                                    totaltime += differ1;
                                                    totalhrs += differ1;
                                                    totaltime1 += differ1;
                                                }
                                            }

                                            if (j == 3)
                                            {
                                                dt2.Rows[cnt + m]["lunchout"] = dt3.Rows[j]["intime"].ToString();
                                                dt2.Rows[cnt + m]["idcardno"] = dt1.Rows[i]["idcardno"].ToString();
                                                dt2.Rows[cnt + m]["empname"] = dt1.Rows[i]["empname"].ToString();

                                            }

                                            if (j == 4)
                                            {
                                                dt2.Rows[cnt + m]["lucnchin"] = dt3.Rows[j]["intime"].ToString();
                                                dt2.Rows[cnt + m]["idcardno"] = dt1.Rows[i]["idcardno"].ToString();
                                                dt2.Rows[cnt + m]["empname"] = dt1.Rows[i]["empname"].ToString();
                                                if (dt2.Rows[cnt + m]["lucnchin"].ToString() != "")
                                                {
                                                    DateTime statetime = DateTime.Parse(dt2.Rows[cnt + m]["lunchout"].ToString());
                                                    DateTime endtime = DateTime.Parse(dt2.Rows[cnt + m]["lucnchin"].ToString());
                                                    TimeSpan differ = endtime.Subtract(statetime);
                                                    TimeSpan differ1 = endtime - statetime;
                                                    dt2.Rows[cnt + m]["luncexcess"] = differ1.ToString();
                                                    totaltime += differ1;
                                                    totalhrs += differ1;
                                                    totaltime1 += differ1;
                                                }
                                            }

                                            if (j == 5)
                                            {
                                                dt2.Rows[cnt + m]["eveteaout"] = dt3.Rows[j]["intime"].ToString();
                                                dt2.Rows[cnt + m]["idcardno"] = dt1.Rows[i]["idcardno"].ToString();
                                                dt2.Rows[cnt + m]["empname"] = dt1.Rows[i]["empname"].ToString();
                                            }

                                            if (j == 6)
                                            {
                                                dt2.Rows[cnt + m]["eveteain"] = dt3.Rows[j]["intime"].ToString();
                                                dt2.Rows[cnt + m]["idcardno"] = dt1.Rows[i]["idcardno"].ToString();
                                                dt2.Rows[cnt + m]["empname"] = dt1.Rows[i]["empname"].ToString();
                                                if (dt2.Rows[cnt + m]["eveteain"].ToString() != "")
                                                {
                                                    DateTime statetime = DateTime.Parse(dt2.Rows[cnt + m]["eveteaout"].ToString());
                                                    DateTime endtime = DateTime.Parse(dt2.Rows[cnt + m]["eveteain"].ToString());
                                                    TimeSpan differ = endtime.Subtract(statetime);
                                                    TimeSpan differ1 = endtime - statetime;
                                                    dt2.Rows[cnt + m]["eventeaexcess"] = differ1.ToString();
                                                    totaltime += differ1;
                                                    totalhrs += differ1;
                                                    dt2.Rows[cnt + m]["totexcess"] = totaltime.ToString();
                                                    totaltime1 += differ1;
                                                }
                                            }

                                            if (j == 7)
                                            {

                                                dt2.Rows[cnt + m]["per1"] = dt3.Rows[j]["intime"].ToString();
                                                dt2.Rows[cnt + m]["idcardno"] = dt1.Rows[i]["idcardno"].ToString();
                                                dt2.Rows[cnt + m]["empname"] = dt1.Rows[i]["empname"].ToString();

                                            }

                                            if (j == 8)
                                            {
                                                dt2.Rows[cnt + m]["per2"] = dt3.Rows[j]["intime"].ToString();
                                                dt2.Rows[cnt + m]["idcardno"] = dt1.Rows[i]["idcardno"].ToString();
                                                dt2.Rows[cnt + m]["empname"] = dt1.Rows[i]["empname"].ToString();
                                                if (dt2.Rows[cnt + m]["per2"].ToString() != "")
                                                {
                                                    DateTime statetime = DateTime.Parse(dt2.Rows[cnt + m]["per1"].ToString());
                                                    DateTime endtime = DateTime.Parse(dt2.Rows[cnt + m]["per2"].ToString());
                                                    TimeSpan differ = endtime.Subtract(statetime);
                                                    TimeSpan differ1 = endtime - statetime;
                                                    dt2.Rows[cnt + m]["per3"] = differ1.ToString();
                                                    totaltime += differ1;
                                                    totalhrs += differ1;
                                                    dt2.Rows[cnt + m]["totexcess"] = totaltime.ToString();
                                                    totaltime1 += differ1;
                                                }
                                            }

                                            if (j == 9)
                                            {
                                                dt2.Rows[cnt + m]["per4"] = dt3.Rows[j]["intime"].ToString();
                                                dt2.Rows[cnt + m]["idcardno"] = dt1.Rows[i]["idcardno"].ToString();
                                                dt2.Rows[cnt + m]["empname"] = dt1.Rows[i]["empname"].ToString();
                                            }

                                            if (j == 10)
                                            {
                                                dt2.Rows[cnt + m]["per5"] = dt3.Rows[j]["intime"].ToString();
                                                dt2.Rows[cnt + m]["idcardno"] = dt1.Rows[i]["idcardno"].ToString();
                                                dt2.Rows[cnt + m]["empname"] = dt1.Rows[i]["empname"].ToString();
                                                if (dt2.Rows[cnt + m]["eveteain"].ToString() != "")
                                                {
                                                    DateTime statetime = DateTime.Parse(dt2.Rows[cnt + m]["per4"].ToString());
                                                    DateTime endtime = DateTime.Parse(dt2.Rows[cnt + m]["per5"].ToString());
                                                    TimeSpan differ = endtime.Subtract(statetime);
                                                    TimeSpan differ1 = endtime - statetime;
                                                    dt2.Rows[cnt + m]["per6"] = differ1.ToString();
                                                    totaltime += differ1;
                                                    totalhrs += differ1;
                                                    dt2.Rows[cnt + m]["totexcess"] = totaltime.ToString();
                                                    totaltime1 += differ1;
                                                }
                                            }

                                            if (j == 11)
                                            {
                                                dt2.Rows[cnt + m]["per6"] = dt3.Rows[j]["intime"].ToString();
                                                dt2.Rows[cnt + m]["idcardno"] = dt1.Rows[i]["idcardno"].ToString();
                                                dt2.Rows[cnt + m]["empname"] = dt1.Rows[i]["empname"].ToString();

                                            }

                                            if (j == 12)
                                            {
                                                dt2.Rows[cnt + m]["per7"] = dt3.Rows[j]["intime"].ToString();
                                                dt2.Rows[cnt + m]["idcardno"] = dt1.Rows[i]["idcardno"].ToString();
                                                dt2.Rows[cnt + m]["empname"] = dt1.Rows[i]["empname"].ToString();
                                                if (dt2.Rows[cnt + m]["eveteain"].ToString() != "")
                                                {
                                                    DateTime statetime = DateTime.Parse(dt2.Rows[cnt + m]["per6"].ToString());
                                                    DateTime endtime = DateTime.Parse(dt2.Rows[cnt + m]["per7"].ToString());
                                                    TimeSpan differ = endtime.Subtract(statetime);
                                                    TimeSpan differ1 = endtime - statetime;
                                                    dt2.Rows[cnt + m]["per9"] = differ1.ToString();
                                                    totaltime += differ1;
                                                    totalhrs += differ1;
                                                    dt2.Rows[cnt + m]["totexcess"] = totaltime.ToString();
                                                    totaltime1 += differ1;
                                                }
                                            }
                                            if (j == 0 && totalcount3 == 1 || j == 1 && totalcount3 == 1 || j == 2 && totalcount3 == 1 || j == 3 && totalcount3 == 1 || j == 4 && totalcount3 == 1 || j == 5 && totalcount3 == 1 || j == 6 && totalcount3 == 1 || j == 7 && totalcount3 == 1 || j == 8 && totalcount3 == 1 || j == 9 && totalcount3 == 1 || j == 10 && totalcount3 == 1 || j == 11 && totalcount3 == 1 || j == 12 && totalcount3 == 1 || j == 13)
                                            {
                                                dt2.Rows[cnt + m]["idcardno"] = dt1.Rows[i]["idcardno"].ToString();
                                                dt2.Rows[cnt + m]["empname"] = dt1.Rows[i]["empname"].ToString();
                                                val1 = ""; TimeSpan differ;
                                                if (dt2.Rows[cnt + m]["minimumin"].ToString() != "")
                                                {
                                                    string sel7 = "SELECT max(a.outtime) as intime,c.PERMISSIONHRS  FROM " + Class.Users.HCompcode + "TRS_ATTLOG a join asptblemp b on a.enrollno=b.idcardno join gtcompmast c on c.gtcompmastid = b.compcode where c.compcode='" + Class.Users.HCompcode + "' and  b.idcardno ='" + dt1.Rows[i]["idcardno"].ToString() + "' and a.indate='" + Convert.ToDateTime(dt1.Rows[i]["indate"].ToString()).ToString("yyyy-MM-dd") + "' group by c.PERMISSIONHRS order by 1 ";
                                                    DataSet ds7 = Utility.ExecuteSelectQuery(sel7, "ASPTBLEMP");
                                                    DataTable dt7 = ds7.Tables["ASPTBLEMP"];
                                                    TimeSpan minimumin = TimeSpan.Parse(dt2.Rows[cnt + m]["minimumin"].ToString());
                                                    TimeSpan mintime = TimeSpan.Parse(dt2.Rows[cnt + m]["starttime"].ToString());//minimumindt2.Rows[cnt + m]["minimumin"].ToString()
                                                    TimeSpan maxtime = TimeSpan.Parse(dt7.Rows[0]["intime"].ToString());
                                                    TimeSpan diff = minimumin.Subtract(mintime);
                                                    differ = maxtime.Subtract(mintime);
                                                    TimeSpan offstarttime1; TimeSpan offstarttime2;
                                                    if (diff.Hours > 0 || diff.Minutes > 0 || diff.Seconds > 0)
                                                    {
                                                        differ = maxtime.Subtract(minimumin);
                                                        offstarttime1 = offtime.Subtract(minimumin);
                                                        offstarttime2 = maxtime.Subtract(offentime);
                                                    }
                                                    else
                                                    {
                                                        differ = maxtime.Subtract(mintime);
                                                        offstarttime1 = offtime.Subtract(mintime);
                                                        offstarttime2 = maxtime.Subtract(offentime);
                                                    }

                                                    val1 = offstarttime1.ToString().Substring(0, 1);
                                                    if (combotype.Text == "Attendance")
                                                    {
                                                        da = 0; ho = 0; mi = 0; se = 0; totmits = 0; totmits1 = 0; totmits2 = 0;
                                                        da = offstarttime1.Days;
                                                        ho = offstarttime1.Hours * 60;
                                                        mi = offstarttime1.Minutes;
                                                        totmits = da + ho + mi;
                                                        totmits1 = totmits;
                                                        da = 0; ho = 0; mi = 0; se = 0; totmits = 0;
                                                        da = offstarttime2.Days;
                                                        ho = offstarttime2.Hours * 60;
                                                        mi = offstarttime2.Minutes;
                                                        totmits = da + ho + mi;
                                                        totmits2 = totmits;
                                                        if (offstarttime1.Minutes <= 0)
                                                        {

                                                            val1 = totmits1.ToString().Trim('-');
                                                            TimeSpan result1 = TimeSpan.FromMinutes(Convert.ToInt64(val1));
                                                            da = 0; ho = 0; mi = 0; se = 0; totmits = 0; totmits1 = 0; totmits1 = 0;
                                                            da = result1.Days;
                                                            ho = result1.Hours * 60;
                                                            mi = result1.Minutes;
                                                            totmits = da + ho + mi; totmits1 = 0;
                                                            totmits1 = totmits;
                                                            da = 0; ho = 0; mi = 0; se = 0; totmits = 0; val1 = "";
                                                            if (offstarttime2.Ticks > 0)
                                                            {
                                                                totmits2 = 0; totmits = 0;
                                                            }
                                                            else
                                                            {
                                                                val1 = totmits2.ToString().Trim('-');
                                                                TimeSpan result2 = TimeSpan.FromMinutes(Convert.ToInt64(val1));
                                                                da = result2.Days;
                                                                ho = result2.Hours * 60;
                                                                mi = result2.Minutes;
                                                                totmits = da + ho + mi; totmits2 = 0;
                                                                totmits2 = totmits;
                                                            }
                                                            val1 = "";


                                                        }


                                                        Int64 tiemdiff = totmits2 + totmits1;

                                                        val1 = tiemdiff.ToString().Trim('-');
                                                        TimeSpan result = TimeSpan.FromMinutes(Convert.ToInt64(val1));
                                                        pertime = TimeSpan.Zero; ho = 0; mi = 0;

                                                        if (dtl.Rows.Count > 0 && dtl.Rows[0]["permission"].ToString() != "")
                                                        {
                                                            pertime = TimeSpan.Parse(dtl.Rows[0]["permission"].ToString());
                                                            ho = pertime.Hours * 60;
                                                            mi = pertime.Minutes;
                                                        }
                                                        string fromTimeString = result.ToString("hh':'mm");

                                                        Int64 totper = ho + mi + tiemdiff;
                                                        dt2.Rows[cnt + m]["per1"] = totper.ToString();
                                                        dt2.Rows[cnt + m]["per10"] = fromTimeString.ToString();
                                                    }

                                                    if (dt7.Rows[0]["intime"].ToString() != "")
                                                    {
                                                        TimeSpan enttime1 = TimeSpan.Parse(dt7.Rows[0]["intime"].ToString());
                                                        if (offentime < enttime1)
                                                        {
                                                            dt2.Rows[cnt + m]["maxiumout"] = dt7.Rows[0]["intime"].ToString();
                                                            TimeSpan statetime = TimeSpan.Parse(dt2.Rows[cnt + m]["starttime"].ToString());//minimumin
                                                            TimeSpan differ1 = offentime - statetime;
                                                            if (dtl.Rows.Count > 0)
                                                            {
                                                                dt2.Rows[cnt + m]["per11"] = dtl.Rows[i]["permission"].ToString(); dtl.Rows.Clear();
                                                            }
                                                            if (dt2.Rows[cnt + m]["per11"].ToString() != "")
                                                            {
                                                                TimeSpan enttime2 = TimeSpan.Parse(dt2.Rows[cnt + m]["per11"].ToString());
                                                                enttime10 = TimeSpan.Parse(dt7.Rows[0]["PERMISSIONHRS"].ToString());

                                                                dt2.Rows[cnt + m]["timediffer"] = differ1 + enttime2;
                                                                dtl.Rows.Clear();
                                                                TimeSpan timediff = TimeSpan.Parse(dt2.Rows[cnt + m]["timediffer"].ToString());

                                                                dt2.Rows[cnt + m]["timediffer"] = timediff.ToString().Substring(0, 5);
                                                                dt2.Rows[cnt + m]["per12"] = timediff;
                                                                TimeSpan timediff12 = TimeSpan.Parse(dt2.Rows[cnt + m]["per12"].ToString());
                                                                if (timediff12.Hours >= 1)
                                                                {
                                                                    da = 0; ho = 0; mi = 0; se = 0; totmits = 0; totmits1 = 0; totmits2 = 0;
                                                                    da = timediff12.Days;
                                                                    ho = timediff12.Hours * 60;
                                                                    mi = timediff12.Minutes;

                                                                    totmits = da + ho + mi;
                                                                    dt2.Rows[cnt + m]["per12"] = totmits.ToString();



                                                                }
                                                                else
                                                                {
                                                                    dt2.Rows[cnt + m]["per12"] = timediff12.ToString();


                                                                }


                                                            }
                                                            else
                                                            {

                                                                dt2.Rows[cnt + m]["timediffer"] = differ1;
                                                                enttime10 = TimeSpan.Parse(dt7.Rows[0]["PERMISSIONHRS"].ToString());
                                                                TimeSpan timediff = TimeSpan.Parse(dt2.Rows[cnt + m]["timediffer"].ToString());

                                                                dt2.Rows[cnt + m]["timediffer"] = timediff.ToString().Substring(0, 5);
                                                                dt2.Rows[cnt + m]["per12"] = timediff;

                                                                TimeSpan timediff12 = TimeSpan.Parse(dt2.Rows[cnt + m]["per12"].ToString());
                                                                if (timediff12.Hours >= 1)
                                                                {
                                                                    da = 0; ho = 0; mi = 0; se = 0; totmits = 0; totmits1 = 0; totmits2 = 0;
                                                                    da = timediff12.Days;
                                                                    ho = timediff12.Hours * 60;
                                                                    mi = timediff12.Minutes;

                                                                    totmits = da + ho + mi;
                                                                    dt2.Rows[cnt + m]["per12"] = totmits.ToString();



                                                                }
                                                                else
                                                                {
                                                                    dt2.Rows[cnt + m]["per12"] = timediff12.ToString();


                                                                }



                                                            }
                                                            dt2.Rows[cnt + m]["totexcess"] = totaltime.ToString();
                                                        }
                                                        else
                                                        {


                                                            dt2.Rows[cnt + m]["maxiumout"] = dt7.Rows[0]["intime"].ToString();
                                                            if (dtl.Rows.Count > 0)
                                                            {
                                                                dt2.Rows[cnt + m]["per11"] = dtl.Rows[0]["permission"].ToString(); dtl.Rows.Clear();
                                                            }
                                                            if (dt2.Rows[cnt + m]["per11"].ToString() != "")
                                                            {
                                                                TimeSpan enttime2 = TimeSpan.Parse(dt2.Rows[cnt + m]["per11"].ToString());
                                                                enttime10 = TimeSpan.Parse(dt7.Rows[0]["PERMISSIONHRS"].ToString());
                                                                dt2.Rows[cnt + m]["timediffer"] = differ;// + enttime2;
                                                                dtl.Rows.Clear();
                                                                TimeSpan timediff = TimeSpan.Parse(dt2.Rows[cnt + m]["timediffer"].ToString());
                                                                dt2.Rows[cnt + m]["timediffer"] = timediff.ToString().Substring(0, 5);
                                                                dt2.Rows[cnt + m]["per12"] = timediff;// - enttime10;
                                                                TimeSpan timediff12 = TimeSpan.Parse(dt2.Rows[cnt + m]["per12"].ToString());
                                                                if (timediff12.Hours >= 1)
                                                                {
                                                                    da = 0; ho = 0; mi = 0; se = 0; totmits = 0; totmits1 = 0; totmits2 = 0;
                                                                    da = timediff12.Days;
                                                                    ho = timediff12.Hours * 60;
                                                                    mi = timediff12.Minutes;

                                                                    totmits = da + ho + mi;
                                                                    dt2.Rows[cnt + m]["per12"] = totmits.ToString();



                                                                }
                                                                else
                                                                {
                                                                    if (timediff12.Hours < 0)
                                                                    {
                                                                        dt2.Rows[cnt + m]["per12"] = timediff12.ToString().Substring(0, 6);
                                                                        da = 0; ho = 0; mi = 0; se = 0; totmits = 0; totmits1 = 0;
                                                                        dt2.Rows[cnt + m]["per12"] = timediff12.ToString().Substring(0, 6);
                                                                        da = timediff12.Days;
                                                                        ho = timediff12.Hours * 60;
                                                                        mi = timediff12.Minutes;
                                                                        totmits = da + ho + mi;
                                                                        totmits1 = totmits;
                                                                        dt2.Rows[cnt + m]["per12"] = totmits1.ToString();
                                                                    }
                                                                    else
                                                                    {
                                                                        dt2.Rows[cnt + m]["per12"] = timediff12.ToString().Substring(0, 5);
                                                                    }


                                                                }
                                                            }
                                                            else
                                                            {
                                                                dt2.Rows[cnt + m]["timediffer"] = differ;
                                                                enttime10 = TimeSpan.Parse(dt7.Rows[0]["PERMISSIONHRS"].ToString());
                                                                TimeSpan timediff = TimeSpan.Parse(dt2.Rows[cnt + m]["timediffer"].ToString());
                                                                dt2.Rows[cnt + m]["timediffer"] = timediff.ToString().Substring(0, 5);

                                                                dt2.Rows[cnt + m]["per12"] = timediff;// - enttime10;
                                                                TimeSpan timediff12 = TimeSpan.Parse(dt2.Rows[cnt + m]["per12"].ToString());



                                                                if (timediff12.Hours >= 1)
                                                                {
                                                                    da = 0; ho = 0; mi = 0; se = 0; totmits = 0; totmits1 = 0; totmits2 = 0;
                                                                    da = timediff12.Days;
                                                                    ho = timediff12.Hours * 60;
                                                                    mi = timediff12.Minutes;
                                                                    totmits = da + ho + mi;
                                                                    dt2.Rows[cnt + m]["per12"] = totmits.ToString();


                                                                }
                                                                else
                                                                {
                                                                    if (timediff12.Hours < 0)
                                                                    {
                                                                        da = 0; ho = 0; mi = 0; se = 0; totmits = 0; totmits1 = 0;
                                                                        dt2.Rows[cnt + m]["per12"] = timediff12.ToString().Substring(0, 6);
                                                                        da = timediff12.Days;
                                                                        ho = timediff12.Hours * 60;
                                                                        mi = timediff12.Minutes;
                                                                        totmits = da + ho + mi;
                                                                        totmits1 = totmits;
                                                                        dt2.Rows[cnt + m]["per12"] = totmits1.ToString();


                                                                    }
                                                                    else
                                                                    {
                                                                        dt2.Rows[cnt + m]["per12"] = timediff12.ToString().Substring(0, 5);

                                                                    }


                                                                }

                                                            }
                                                            if (dt2.Rows[cnt + m]["mteain"].ToString() == "")
                                                            {


                                                            }
                                                            else
                                                            {

                                                                dt2.Rows[cnt + m]["totexcess"] = totaltime.ToString();

                                                            }
                                                        }

                                                    }
                                                }
                                            }



                                            totalcount3--;
                                        }
                                    }
                                    dt2.Rows[m]["totalhrs"] = totalhrs.ToString();
                                    dt2.Rows[m]["permissionhrs"] = permissionhrs.ToString();
                                    TimeSpan rtot = TimeSpan.Parse(dt2.Rows[m]["totalhrs"].ToString());
                                    TimeSpan rper = TimeSpan.Parse(dt2.Rows[m]["permissionhrs"].ToString());
                                    TimeSpan rdiffer1 = rtot - rper;
                                    if (totaltime1.Ticks > 0)
                                    {
                                        dt2.Rows[m]["datetimerecord"] = totaltime1.ToString();
                                    }
                                    dt2.Rows[m]["outtime"] = rdiffer1.ToString();
                                    totaltime1 = TimeSpan.Zero;
                                }
                            }
                            else
                            {
                                if (dt2.Rows.Count <= 0)
                                {
                                    dt2 = dt1.Clone();
                                    cnt = 0;
                                }
                                dt2.Rows.Add();
                                sno += 1;
                                dt2.Rows[cnt + m]["asptblempid"] = sno;
                                if (com.ToString() == "")
                                {
                                    com = Class.Users.HCompName;
                                    add = "Address";
                                    idcard = comboidcardno.Text;
                                    empname = txtempname.Text;
                                }
                                dt2.Rows[cnt + m]["compname"] = com.ToString();
                                dt2.Rows[cnt + m]["address"] = add.ToString();
                                dt2.Rows[cnt + m]["idcardno"] = idcard.ToString();
                                dt2.Rows[cnt + m]["empname"] = empname.ToString();

                                dt2.Rows[cnt + m]["indate"] = Convert.ToDateTime(dt0.Rows[m]["indate"].ToString()).ToString("dd-MM-yyyy");
                                dt2.Rows[cnt + m]["minimumin"] = "A";
                            }
                        }

                    }
                    else
                    {
                        MessageBox.Show("Invalid");
                    }
                }
                if (checkBox1.Checked == false)
                {
                    if (dt2.Rows.Count >= 1)
                    {
                        var orderedRows = from row in dt2.AsEnumerable() select row;
                        dtprint = orderedRows.CopyToDataTable(); dt2.Rows.Clear();
                    }
                }
                    if (dtprint.Rows.Count >= 1)
                    {

                        string sel0 = "SELECT  CONCAT(a.intime, '     ',a.outtime) as officetime , CONCAT(a.lunchfrom, '     ',a.lunchto) as lunchtime,CONCAT(a.breakfrom, '     ',a.breakto,'          ' ,a.eventeafrom, '    ',a.eventeato) as teatime  from gtcompmast a where a.compcode='" + Class.Users.HCompcode + "'";
                        DataSet ds0 = Utility.ExecuteSelectQuery(sel0, "asptblemp");
                        DataTable dt10 = ds0.Tables["asptblemp"];
                        dtmonth.Rows[0]["officetime"] = dt10.Rows[0]["officetime"].ToString();
                        dtmonth.Rows[0]["lunchtime"] = dt10.Rows[0]["lunchtime"].ToString();
                        dtmonth.Rows[0]["teatime"] = dt10.Rows[0]["teatime"].ToString();
                        dtmonth.Rows[0]["teatime1"] = "From : " + frmdate.Value.ToString().Substring(0, 10) + " To: " + todate.Value.ToString().Substring(0, 10);
                        if (combotype.Text == "Attendance")
                        {
                            tabControl2.SelectTab(tabPage3);
                            rd1.Database.Tables["DataTable11"].SetDataSource(dtprint);
                            rd1.Database.Tables["DataTable2"].SetDataSource(dtmonth);
                            crystalReportViewer2.ReportSource = null;
                            crystalReportViewer2.ReportSource = rd1;
                            crystalReportViewer2.Refresh(); comboidcardno.Select(); dtprint.Rows.Clear();
                            crystalReportViewer2.Zoom(150);
                        }
                        else
                        {
                            tabControl2.SelectTab(tabPage2);
                            rd.Database.Tables["DataTable1"].SetDataSource(dtprint);
                            rd.Database.Tables["DataTable2"].SetDataSource(dtmonth);
                            crystalReportViewer1.ReportSource = null;
                            crystalReportViewer1.ReportSource = rd;
                            crystalReportViewer1.Refresh(); comboidcardno.Select(); dtprint.Rows.Clear(); Cursor = Cursors.Default;
                            crystalReportViewer1.Zoom(150);
                        }
                    }
                    Cursor = Cursors.Default;
                }
         
            catch (Exception ex)
            {
                crystalReportViewer1.ReportSource = null;
                crystalReportViewer1.Refresh(); Cursor = Cursors.Default;
            }
            Cursor = Cursors.Default;
        }
        string monthname(string mon)
        {
            switch (mon)
            {
                case "01":
                    mon = "JANUARY";
                    break;
                case "02":
                    mon = "FEBRUARY";
                    break;
                case "03":
                    mon = "MARCH";
                    break;
                case "04":
                    mon = "APRIL";
                    break;
                case "05":
                    mon = "MAY";
                    break;
                case "06":
                    mon = "JUNE";
                    break;
                case "07":
                    mon = "JULY";
                    break;
                case "08":
                    mon = "AUGUST";
                    break;
                case "09":
                    mon = "SEPTEMBER";
                    break;
                case "10":
                    mon = "OCTOBER";
                    break;
                case "11":
                    mon = "NOVEMBER";
                    break;
                case "12":
                    mon = "DECEMBER";
                    break;
            }
            return mon;
        }
        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {

        }

        public void News()
        {
            dtprint.Rows.Clear(); dt2.Rows.Clear();crystalReportViewer1.Refresh();crystalReportViewer1.ReportSource = null;
            crystalReportViewer2.Refresh(); crystalReportViewer2.ReportSource = null;
            permissionhrs = TimeSpan.Zero; remainhrs = TimeSpan.Zero; totaltime = TimeSpan.Zero; totalhrs = TimeSpan.Zero;
            panel1.BackColor = Class.Users.BackColors;          
            this.Font= Class.Users.FontName;          Class.Users.UserTime = 0;
            crystalReportViewer1.Font= Class.Users.FontName;
            //DateTime selectedDate = MonthYearDT();
            //DateTime selectedDate1 = MonthYearDT1();
            //var lastDay = DateTime.DaysInMonth(System.DateTime.Now.Year, System.DateTime.Now.Month);
            //var first = new DateTime(System.DateTime.Now.Year, System.DateTime.Now.Month, 1);
            //var first1 = new DateTime(System.DateTime.Now.Year, System.DateTime.Now.Month, lastDay);
            //frmdate.Value = first;
            //todate.Value = first1;
        }

        public void Saves()
        {
           
        }

        public void Prints()
        {
            
        }

        public void Searchs()
        {
            
        }

        public void Deletes()
        {
            
        }

        public void ReadOnlys()
        {
            
        }

        public void Imports()
        {
           
        }

        public void Pdfs()
        {
           
        }

        public void ChangePasswords()
        {
            
        }

        public void DownLoads()
        {
            if (comboformate.Text != "")
            {

                DialogResult result = MessageBox.Show("Do you want to '" + comboformate.Text + "' Formate ??", "" + comboformate.Text + "DownLoad", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                if (result.Equals(DialogResult.OK))
                {
                    if (checkBox1.Checked == true)
                    {
                        switch (comboformate.Text)
                        {
                            case "Word":
                                if (combotype.Text == "Attendance")
                                {
                                    rd1.ExportToDisk(ExportFormatType.WordForWindows, "E:\\'" + System.DateTime.Now.ToString("dd-MM-yyyy") + "  ALL -Attendance.doc");
                                }
                                else
                                {
                                    rd.ExportToDisk(ExportFormatType.WordForWindows, "E:\\'" + System.DateTime.Now.ToString("dd-MM-yyyy") + "  ALL -DateTimeReport.doc");
                                }
                                break;

                            case "Excel":
                                if (combotype.Text == "Attendance")
                                {
                                    rd1.ExportToDisk(ExportFormatType.ExcelRecord, "E:\\'" + System.DateTime.Now.ToString("dd-MM-yyyy") + "  ALL -Attendance.xls");
                                }
                                else
                                {
                                    rd.ExportToDisk(ExportFormatType.ExcelRecord, "E:\\'" + System.DateTime.Now.ToString("dd-MM-yyyy") + "  ALL -DateTimeReport.xls");
                                }
                                break;

                            case "PDF":
                                if (combotype.Text == "Attendance")
                                {
                                    rd1.ExportToDisk(ExportFormatType.PortableDocFormat, "E:\\'" + System.DateTime.Now.ToString("dd-MM-yyyy") + "  ALL -Attendance.pdf");
                                }
                                else
                                {
                                    rd.ExportToDisk(ExportFormatType.PortableDocFormat, "E:\\'" + System.DateTime.Now.ToString("dd-MM-yyyy") + "  ALL -DateTimeReport.pdf");
                                }
                                break;

                            case "CSV":
                                if (combotype.Text == "Attendance")
                                {
                                    rd1.ExportToDisk(ExportFormatType.CharacterSeparatedValues, "E:\\'" + System.DateTime.Now.ToString("dd-MM-yyyy") + "  ALL -Attendance.csv");
                                }
                                else
                                {
                                    rd.ExportToDisk(ExportFormatType.CharacterSeparatedValues, "E:\\'" + System.DateTime.Now.ToString("dd-MM-yyyy") + "  ALL -DateTimeReport.csv");
                                }
                                break;
                        }
                        checkBox1.Checked = false;
                    }
                    else
                    {
                        switch (comboformate.Text)
                        {
                            case "Word":
                                if (combotype.Text == "Attendance")
                                {
                                    rd1.ExportToDisk(ExportFormatType.WordForWindows, "E:\\'" + System.DateTime.Now.ToString("dd-MM-yyyy") + "  " + comboidcardno.Text + "'Attendance.doc");
                                }
                                else
                                {
                                    rd.ExportToDisk(ExportFormatType.WordForWindows, "E:\\'" + System.DateTime.Now.ToString("dd-MM-yyyy") + "  " + comboidcardno.Text + "'DateTimeReport.doc");
                                }
                                break;

                            case "Excel":
                                if (combotype.Text == "Attendance")
                                {
                                    rd1.ExportToDisk(ExportFormatType.ExcelRecord, "E:\\'" + System.DateTime.Now.ToString("dd-MM-yyyy") + "  " + comboidcardno.Text + "'Attendance.xls");
                                }
                                else
                                {
                                    rd.ExportToDisk(ExportFormatType.ExcelRecord, "E:\\'" + System.DateTime.Now.ToString("dd-MM-yyyy") + "  " + comboidcardno.Text + "'DateTimeReport.xls");
                                }
                                break;

                            case "PDF":
                                if (combotype.Text == "Attendance")
                                {
                                    rd1.ExportToDisk(ExportFormatType.PortableDocFormat, "E:\\'" + System.DateTime.Now.ToString("dd-MM-yyyy") + "  " + comboidcardno.Text + "'Attendance.pdf");
                                }
                                else
                                {
                                    rd.ExportToDisk(ExportFormatType.PortableDocFormat, "E:\\'" + System.DateTime.Now.ToString("dd-MM-yyyy") + "  " + comboidcardno.Text + "'DateTimeReport.pdf");
                                }
                                break;

                            case "CSV":
                                if (combotype.Text == "Attendance")
                                {
                                    rd1.ExportToDisk(ExportFormatType.CharacterSeparatedValues, "E:\\'" + System.DateTime.Now.ToString("dd-MM-yyyy") + "  " + comboidcardno.Text + "'Attendance.csv");
                                }
                                else
                                {
                                    rd.ExportToDisk(ExportFormatType.CharacterSeparatedValues, "E:\\'" + System.DateTime.Now.ToString("dd-MM-yyyy") + "  " + comboidcardno.Text + "'DateTimeReport.csv");
                                }
                                break;
                        }
                    }


                }
                else
                {

                }
            }
            else
            {
                MessageBox.Show("Pls Select Combo Box Value");
            }
        }

        public void ChangeSkins()
        {
            throw new NotImplementedException();
        }

        public void Logins()
        {
            throw new NotImplementedException();
        }

        public void GlobalSearchs()
        {
            throw new NotImplementedException();
        }

        public void TreeButtons()
        {
            throw new NotImplementedException();
        }

        public void Exit()
        {
            GlobalVariables.MdiPanel.Show();
            this.Hide();
            News();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);
        }

        public void GridLoad()
        {
           
        }

        private void refresgToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DateTimeReport_Load(sender,e);
        }

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }
        private void tabControl2_SelectedIndexChanged(object sender, EventArgs e)
        {
         
            if (tabControl2.SelectedTab == tabControl2.TabPages["tabPage4"])//your specific tabname
            {
                try
                {
                    DataTable dt1 = com.select("select distinct e.compname,e.address,d.empname,b.enrollno as idcardno,'' idcard0,'' idcard1,''idcard2,'' idcard3,'' idcard4,'' idcard5,'' idcard6,'' idcard7,''idcard8,'' idcard9,'' idcard10,'' idcard11,'' idcard12,'' idcard13,'' idcard14,'' idcard15,'' idcard16,'' idcard17,'' idcard18,'' idcard19,'' idcard20,'' idcard21,'' idcard22,'' idcard23,'' idcard24,'' idcard25,'' idcard26,'' idcard27,'' idcard28,'' idcard29,'' idcard30,'' idcard31,'' date0,'' date1,''date2,'' date3,'' date4,'' date5,'' date6,'' date7,''date8,'' date9,'' date10,'' date11,'' date12,'' date13,'' date14,'' date15,'' date16,'' date17,'' date18,'' date19,'' date20,'' date21,'' date22,'' date23,'' date24,'' date25,'' date26,'' date27,'' date28,'' date29,'' date30,'' in0,'' in1,''in2,'' in3,'' in4,'' in5,'' in6,'' in7,''in8,'' in9,'' in10,'' in11,'' in12,'' in13,'' in14,'' in15,'' in16,'' in17,'' in18,'' in19,'' in20,'' in21,'' in22,'' in23,'' in24,'' in25,'' in26,'' in27,'' in28,'' in29,'' in30,'' in31 from pinTRS_ATTLOG b join asptblemp d on  d.idcardno=b.enrollno  join gtcompmast e on e.gtcompmastid=d.compcode where d.active='T'  and  b.compcode='" + Class.Users.HCompcode + "' order by 4", "pinTRS_ATTLOG");
                    DataTable dt2 = com.select("select date_format(indate, '%d-%m-%y') as indate from(select adddate('1970-01-01', t4.i * 10000 + t3.i * 1000 + t2.i * 100 + t1.i * 10 + t0.i) indate from" +
         "(select 0 i union select 1 union select 2 union select 3 union select 4 union select 5 union select 6 union select 7 union select 8 union select 9) t0," +
         "(select 0 i union select 1 union select 2 union select 3 union select 4 union select 5 union select 6 union select 7 union select 8 union select 9) t1," +
         "(select 0 i union select 1 union select 2 union select 3 union select 4 union select 5 union select 6 union select 7 union select 8 union select 9) t2," +
         "(select 0 i union select 1 union select 2 union select 3 union select 4 union select 5 union select 6 union select 7 union select 8 union select 9) t3," +
         "(select 0 i union select 1 union select 2 union select 3 union select 4 union select 5 union select 6 union select 7 union select 8 union select 9) t4) ss " +
         " where indate between '" + frmdate.Value.ToString("yyyy-MM-dd") + "' and '" + todate.Value.ToString("yyyy-MM-dd") + "' order by 1", "dual");
                    int i = 0; int j = 0, k = 0, l = 0; int cnt = 0; int cnt1 = 0;

                    int dat = 0;
                    for (i = 0; i < dt1.Rows.Count; i++)
                    {
                        cnt = i; 
                        for (j = 0; j < dt2.Rows.Count; j++)
                        {
                            dat = Convert.ToInt32("0" + dt2.Rows[j]["indate"].ToString().Substring(0, 2));
                            DataTable dt3 = com.select("select distinct b.enrollno, date_format(b.indate, '%d-%m-%y') as indate from pinTRS_ATTLOG b join asptblemp d on  d.idcardno=b.enrollno where d.active='T' and  date_format(b.indate, '%d-%m-%y') = '" + dt2.Rows[j]["indate"].ToString() + "' and  b.enrollno ='" + dt1.Rows[i]["idcardno"].ToString() + "'    order by 1", "pinTRS_ATTLOG");// and date_format(b.indate, '%d-%m-%y') = '01-07-22' 
                            if (dt3.Rows.Count > 0)
                            {
                                for (k = 0; k < dt3.Rows.Count; k++)
                                {
                                    DataTable dt4 = com.select("select distinct min(b.intime) as intime,max(b.outtime) as outtime, b.enrollno,date_format(b.indate, '%d-%m-%y') as indate from pinTRS_ATTLOG b join asptblemp c on c.idcardno=b.enrollno where  date_format(b.indate, '%d-%m-%y') = '" + dt3.Rows[k]["indate"].ToString() + "' and   b.enrollno ='" + dt1.Rows[i]["idcardno"].ToString() + "'  group by b.enrollno,b.indate   order by 4", "pinTRS_ATTLOG");
                                   dat = Convert.ToInt32("0"+dt4.Rows[k]["indate"].ToString().Substring(0,2));
                                    if (dt4.Rows.Count > 0)
                                    {
                                        for (l = 0; l < dt4.Rows.Count; l++)
                                        {
                                            if (dt1.Rows[cnt]["date" + l + ""].ToString() == "")
                                            {
                                                dt1.Rows[cnt]["date" + l + ""] = dat;// dt4.Rows[l]["indate"].ToString();
                                                dt1.Rows[cnt]["idcard" + l + ""] = dt4.Rows[l]["intime"].ToString();
                                                if (dt4.Rows[l]["intime"].ToString() != dt4.Rows[l]["outtime"].ToString())
                                                {                                               
                                                    dt1.Rows[cnt]["in" + l + ""] = dt4.Rows[l]["outtime"].ToString();
                                                }
                                            }
                                            else
                                            {
                                                cnt1++;

                                                dt1.Rows[cnt]["date" + j + ""] = dat;// dt4.Rows[l]["indate"].ToString();
                                                dt1.Rows[cnt]["idcard" + j + ""] = dt4.Rows[l]["intime"].ToString();
                                                if (dt4.Rows[l]["intime"].ToString() != dt4.Rows[l]["outtime"].ToString())
                                                {
                                                    dt1.Rows[cnt]["in" + j + ""] = dt4.Rows[l]["outtime"].ToString();
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        dt1.Rows[cnt]["date" + k + ""] = dat;// dt2.Rows[k]["indate"].ToString();
                                        dt1.Rows[cnt]["idcard" + k + ""] = "A";
                                        dt1.Rows[cnt]["in" + k + ""] = "A";
                                    }
                                }
                            }
                            else
                            {
                                dt1.Rows[cnt]["date" + j + ""] = dat;//dt2.Rows[j]["indate"].ToString();
                                dt1.Rows[cnt]["idcard" + j + ""] = "A";
                                dt1.Rows[cnt]["in" + j + ""] = "A";
                            }
                        }
                    }
                    rd3.Database.Tables["DataTabledate"].SetDataSource(dt1);
                    crystalReportViewer4.ReportSource = null;
                    crystalReportViewer4.ReportSource = rd3;
                    crystalReportViewer4.Refresh(); crystalReportViewer4.Zoom(100);
                    tabControl2.SelectTab(tabPage4);
                }
                catch (Exception ex) { MessageBox.Show(ex.ToString()); }
            }
        }

        private void tabControl2_TabIndexChanged(object sender, EventArgs e)
        {
           
        }
    }
}

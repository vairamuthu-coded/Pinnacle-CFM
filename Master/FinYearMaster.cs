using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Pinnacle.Master
{
    public partial class FinYearMaster : Form,ToolStripAccess
    {
        private static FinYearMaster _instance;
        Models.Master mas = new Models.Master();
        Models.Employee em = new Models.Employee();
        Models.UserRights sm = new Models.UserRights();
        ListView listfilter = new ListView();
        System.Text.StringBuilder sb = new System.Text.StringBuilder();
        public FinYearMaster()
        {
            InitializeComponent();
            GlobalVariables.HeaderName.Text = Class.Users.ScreenName; GlobalVariables.CurrentForm = this;
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd/MM/yyyy") + " " + System.DateTime.Now.ToLongTimeString());
           
        }

     
        public static FinYearMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new FinYearMaster();
                GlobalVariables.CurrentForm = _instance; return _instance;
            }
        }

       
    
        private void Txttodate_ValueChanged(object sender, EventArgs e)
        {

            DateTime StartingDate = Convert.ToDateTime(txtfromdate.Value.Date.ToString());
            DateTime EndingDate = Convert.ToDateTime(txttodate.Value.Date.ToString());
             TimeSpan countdays  = EndingDate.Subtract(StartingDate);
            int totaldays = Convert.ToInt32(countdays.Days);
            if (totaldays == 366 || totaldays == 365 || totaldays == 364) 
            {
                txttotaldays.Text = totaldays.ToString();
                txtfinyear.Text = StartingDate.ToString("yy") + "-" + EndingDate.ToString("yy");
            }
            else
            {
                txttotaldays.Text = "";
            }
           
        }
        private List<DateTime> GetDateRange(DateTime StartingDate, DateTime EndingDate)
        {
            if (StartingDate > EndingDate)
            {
                return null;
            }
            List<DateTime> rv = new List<DateTime>();
            DateTime tmpDate = StartingDate;
            do
            {
                rv.Add(tmpDate);
                tmpDate = tmpDate.AddDays(1);
            } while (tmpDate <= EndingDate);
            return rv;
        }
        string coid = "";
        Models.CommonClass CC = new Models.CommonClass();
        public void dtpGrid1(string s)
        {
            try
            {
                DataTable dt = CC.select("select a.asptblpartymasid,a.partyname from asptblpartymas a  where a.active='T' and partyname='" + s + "' ", "asptblpartymas");
                coid = "";
                coid = Convert.ToString(dt.Rows[0]["asptblpartymasid"].ToString());


            }
            catch (Exception EX)
            { }
        }
        private void FinYearMaster_Load(object sender, EventArgs e)
        {
           

            Utility.Load_Combo(combo_compcode, "select gtcompmastid,  compcode from  gtcompmast   ORDER BY 2", "gtcompmastid", "compcode");
            Utility.Load_Combo(dtp1, "select a.asptblpartymasid,a.partyname from asptblpartymas a order by  a.asptblpartymasid  desc", "asptblpartymasid", "partyname");

        }


      public  void GridLoad()
        {
            try
            {
                listView1.Items.Clear();
                sb.Append("select a.gtfinancialyearid, a.finyear, a.FROMDATE as fromdate,A.TODATE  as todate,A.CURRENTFINYR,b.compcode,A.TOTALDAYS from gtfinancialyear a join gtcompmast b on a.compcode=b.gtcompmastid  order by a.gtfinancialyearid desc");
                DataSet ds = Utility.ExecuteSelectQuery(sb.ToString(), "gtfinancialyear");
                DataTable dt = ds.Tables["gtfinancialyear"];
                if (dt != null)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["gtfinancialyearid"].ToString());
                        list.SubItems.Add(myRow["FINYEAR"].ToString());
                        list.SubItems.Add(myRow["FROMDATE"].ToString());
                        list.SubItems.Add(myRow["TODATE"].ToString());
                        list.SubItems.Add(myRow["CURRENTFINYR"].ToString());
                        list.SubItems.Add(myRow["COMPCODE"].ToString());
                        list.SubItems.Add(myRow["TOTALDAYS"].ToString());
                        this.listfilter.Items.Add((ListViewItem)list.Clone());
                        if (i % 2 == 0) { list.BackColor = Color.White; } else { list.BackColor = Color.WhiteSmoke; }
                        listView1.Items.Add(list);
                        i++;
                    }
                    lbltotal.Text = "Total Count    :" + listView1.Items.Count;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void TextBox5_TextChanged(object sender, EventArgs e)
        {

        }
        DateTime  StartingDate, EndingDate; TimeSpan countdays;
        private void Txtfromdate_ValueChanged(object sender, EventArgs e)
        {
            StartingDate = Convert.ToDateTime(txtfromdate.Value.Date.ToString());
            EndingDate = Convert.ToDateTime(txttodate.Value.Date.ToString());
            countdays = EndingDate.Subtract(StartingDate);
            int totaldays = Convert.ToInt32(countdays.Days);
            if (totaldays == 364 || totaldays == 365 || totaldays == 366)
            {
                txttotaldays.Text = totaldays.ToString();
                txtfinyear.Text = StartingDate.ToString("yy") + "-" + EndingDate.ToString("yy");
            }
            else
            {
                txttotaldays.Text = "";
            }
        }

        public void Saves()
        {
           
            try
            {
               
                if (txttotaldays.Text == "")
                {
                    MessageBox.Show("'Total Days Empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txttotaldays.Focus(); return;
                }
                if (txtfinyear.Text == "")
                {
                    MessageBox.Show("'FinYear is empty'  Empty not Allowed  ", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtfinyear.Focus(); return;
                }
                foreach(DataGridViewRow row in dataGridView1.Rows)
                {
                   if(row.Cells[4].Value == null)
                    {
                        MessageBox.Show("Starting Date is Emtpy");
                        row.Cells[4].Selected = true;
                        return;
                    }
                    if (row.Cells[5].Value == null)
                    {
                        MessageBox.Show("Ending Date is Emtpy");
                        row.Cells[5].Selected = true;
                        return;
                    }
                }
                if (txttotaldays.Text != "" && txtfinyear.Text != "" && dataGridView1.Rows.Count>1 && dataGridView2.Rows.Count>1 && dataGridView3.Rows.Count>1)
                {
                    string chk = "";
                    if (radiocurrent.Checked == true) { chk = "T"; } else { chk = "F"; radioclosed.Checked = false; }
                    sb.Clear();
                    sb.Append("select gtfinancialyearid    from  gtfinancialyear     WHERE COMPCODE='" + combo_compcode.SelectedValue + "' and FINYEAR='" + txtfinyear.Text + "' AND  CURRENTFINYR='" + chk + "' ");
                    DataSet ds = Utility.ExecuteSelectQuery(sb.ToString(), "gtfinancialyear");
                    DataTable dt = ds.Tables["gtfinancialyear"];
                    if (GlobalVariables.New_Flg==false)
                    {
                        sb.Clear();
                        //sb.Append("INSERT INTO gtfinancialyear(COMPCODE,FINYEAR,FROMDATE,  TODATE,  CURRENTFINYR,  TOTALDAYS,  USERNAME,  MODIFIEDON,  CREATEDON,IPADDRESS)  VALUES('" + combo_compcode.SelectedValue + "','" + txtfinyear.Text + "','" + txtfromdate.Text + "','" + txttodate.Text + "','" + chk + "'," + txttotaldays.Text + "," + Class.Users.USERID + ",'" + Class.Users.CREATED + "','" + Class.Users.CREATED + "','" + Class.Users.IPADDRESS + "' )");
                        //Utility.ExecuteNonQuery(sb.ToString());
                        GridView1(dataGridView1);
                        GridView2(dataGridView2);
                        GridView3(dataGridView3);
                    }
                    if (GlobalVariables.New_Flg == true)
                    {
                        //sb.Clear();
                        //sb.Append("UPDATE  gtfinancialyear  SET COMPCODE='" + combo_compcode.SelectedValue + "'  ,FINYEAR='" + txtfinyear.Text + "' , FROMDATE='" + txtfromdate.Text + "' ,  TODATE='" + txttodate.Text + "' , CURRENTFINYR='" + chk + "' ,TOTALDAYS=" + txttotaldays.Text + " ,USERNAME=" + Class.Users.USERID + ",MODIFIEDON='" + Class.Users.CREATED + "',IpAddress='" + Class.Users.IPADDRESS + "' WHERE gtfinancialyearid=" + txtfinyearid.Text);
                        //Utility.ExecuteNonQuery(sb.ToString());
                        GridView1(dataGridView1);
                        GridView2(dataGridView2);
                        GridView3(dataGridView3);
                    }

                    if (txtfinyearid.Text == "")
                    {
                        Cursor.Current = Cursors.Default;
                        MessageBox.Show("Record Saved Successfully " + txtfinyearid.Text, " Success Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(); empty();


                    }
                    else
                    {
                        MessageBox.Show("Record Updated Successfully " + txtfinyearid.Text, " Update Message ", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        GridLoad(); empty();

                    }
                }

            }
            catch (Exception ex)
            {

                MessageBox.Show("FinYear " + "        " + ex.ToString(), "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }
        string maxid = "";
        Pinnacle.Models.Tally.FinYearModel3 pp = new Models.Tally.FinYearModel3();
       
        private void GridView1(DataGridView dataGridView1)
        {

            int i = 0, j = 1; int cnt = dataGridView1.Rows.Count;
            if (cnt >= 0)
            {
                sb.Clear();sb.Append("select max(gtfinancialyearid) id    from  gtfinancialyear   where  compcode='" + Class.Users.COMPCODE + "'  and finyear='" + txtfinyear.Text.Trim() + "' ");
                DataSet ds2 = Utility.ExecuteSelectQuery(sb.ToString(), "gtfinancialyear");
                DataTable dt2 = ds2.Tables["gtfinancialyear"];
                maxid = dt2.Rows[0]["id"].ToString();

                for (i = 0; i < cnt; i++)
                {
                    if (dataGridView1.Rows[i].Cells[0].Value == null)
                    {
                        dataGridView1.Rows[i].Cells[0].Value = 0;
                    }
                    else
                    {
                        pp.Asptblfindetid = Convert.ToInt64("0" + dataGridView1.Rows[i].Cells[0].Value.ToString());
                    }                  
                    pp.Asptblfindet1id = Convert.ToInt64("0" + maxid);
                    pp.Asptblfinid = Convert.ToInt64("0" + maxid);
                    pp.Finyear = txtfinyear.Text;
                    pp.CompCode = Convert.ToInt64("0" + Class.Users.COMPCODE);
                    //Asptblfindetid,Asptblfinid,Finyear,Compcode,PeriodStartDate,PeriodEndDate,PeriodCode,PeriodDescription,
                    //SalCalDesc,Quarter,TotalDays,Notes
                    pp.PeriodStartDate = Convert.ToDateTime(dataGridView1.Rows[i].Cells[4].EditedFormattedValue.ToString().Substring(0,10));
                    pp.PeriodEndDate = Convert.ToDateTime(dataGridView1.Rows[i].Cells[5].EditedFormattedValue.ToString().Substring(0, 10));
                    pp.PeriodCode = Convert.ToInt64("0" + dataGridView1.Rows[i].Cells[6].Value.ToString().Trim());
                    pp.PeriodDescription = Convert.ToString(dataGridView1.Rows[i].Cells[7].Value.ToString().Trim());
                    pp.SalCalDesc = Convert.ToString(dataGridView1.Rows[i].Cells[8].Value.ToString().Trim());
                    pp.Quarter = Convert.ToString(dataGridView1.Rows[i].Cells[9].Value.ToString().Trim());
                    pp.TotalDays = Convert.ToInt64("0" + dataGridView1.Rows[i].Cells[10].Value.ToString().Trim());
                    dtpGrid1(dataGridView1.Rows[i].Cells[11].Value.ToString().Trim());
                    pp.Notes = coid;

                    sb.Clear();sb.Append("select Asptblfindetid   from  Asptblfindet   where   Asptblfinid='"+pp.Asptblfinid+ "' and Finyear='" + pp.Finyear + "' and Compcode='" + pp.CompCode + "' and PeriodStartDate=date_format('" + pp.PeriodStartDate.ToString("yyyy-MM-dd").Substring(0, 10) + "','%Y-%m-%d') and PeriodEndDate=date_format('" + pp.PeriodEndDate.ToString("yyyy-MM-dd").Substring(0, 10) + "','%Y-%m-%d') and PeriodCode='" + pp.PeriodCode + "' and PeriodDescription='" + pp.PeriodDescription + "' and SalCalDesc='" + pp.SalCalDesc + "' and Quarter='" + pp.Quarter + "' and TotalDays='" + pp.TotalDays + "' and Notes='" + pp.Notes + "' ");
                    DataSet ds1 = Utility.ExecuteSelectQuery(sb.ToString(), "Asptblfindet");
                    DataTable dt1 = ds1.Tables["Asptblfindet"];
                    if (dt1.Rows.Count != 0)
                    {
                       
                    }
                     else if (dt1.Rows.Count != 0 && pp.Asptblfindetid == 0 || pp.Asptblfindetid == 0)
                    {
                        sb.Clear(); sb.Append("insert into Asptblfindet(Asptblfinid,Finyear,Compcode,PeriodStartDate,PeriodEndDate,PeriodCode,PeriodDescription,SalCalDesc,Quarter,TotalDays,Notes) values('" + pp.Asptblfinid + "','" + pp.Finyear + "','" + pp.CompCode + "',date_format('" + pp.PeriodStartDate.ToString("yyyy-MM-dd").Substring(0, 10) + "','%Y-%m-%d'),date_format('" + pp.PeriodEndDate.ToString("yyyy-MM-dd").Substring(0, 10) + "','%Y-%m-%d'),'" + pp.PeriodCode + "' , '" + pp.PeriodDescription + "' ,'" + pp.SalCalDesc + "' , '" + pp.Quarter + "','" + pp.TotalDays + "','" + pp.Notes+ "')");
                        Utility.ExecuteNonQuery(sb.ToString());                                          
                    }
                   else
                    {
                        sb.Clear(); sb.Append("update  Asptblfindet   set   Asptblfinid='" + pp.Asptblfinid + "' , Finyear='" + pp.Finyear + "' , Compcode='" + pp.CompCode + "' , PeriodStartDate=date_format('" + pp.PeriodStartDate.ToString("yyyy-MM-dd").Substring(0, 10) + "','%Y-%m-%d') , PeriodEndDate=date_format('" + pp.PeriodEndDate .ToString("yyyy-MM-dd").Substring(0, 10) + "','%Y-%m-%d') , PeriodCode='" + pp.PeriodCode + "' , PeriodDescription='" + pp.PeriodDescription + "' , SalCalDesc='" + pp.SalCalDesc + "' , Quarter='" + pp.Quarter + "' , TotalDays='" + pp.TotalDays + "' , Notes='" + pp.Notes + "'  where Asptblfindetid='" + pp.Asptblfindetid + "' and Finyear='" + pp.Finyear + "' and Compcode='" + pp.CompCode + "' ");
                        Utility.ExecuteNonQuery(sb.ToString());
                    }
                }
            }
            
        }
        private void GridView2(DataGridView dataGridView1)
        {

            int i = 0, j = 1; int cnt = dataGridView1.Rows.Count;
            if (cnt >= 0)
            {
                sb.Clear(); sb.Append("select max(gtfinancialyearid) id    from  gtfinancialyear   where  compcode='" + Class.Users.COMPCODE + "'  and finyear='" + txtfinyear.Text.Trim() + "' ");
                DataSet ds2 = Utility.ExecuteSelectQuery(sb.ToString(), "gtfinancialyear");
                DataTable dt2 = ds2.Tables["gtfinancialyear"];
                maxid = dt2.Rows[0]["id"].ToString();

                for (i = 0; i < cnt; i++)
                {
                    if (dataGridView1.Rows[i].Cells[0].Value == null)
                    {
                        dataGridView1.Rows[i].Cells[0].Value = 0;
                    }
                    else
                    {
                        pp.asptblfinwekdetid = Convert.ToInt64("0" + dataGridView1.Rows[i].Cells[0].Value.ToString());
                    }
                    pp.asptblfinwekdet1id = Convert.ToInt64("0" + maxid);
                    pp.Asptblfinid = Convert.ToInt64("0" + maxid);
                    pp.Finyear = txtfinyear.Text;
                    pp.CompCode = Convert.ToInt64("0" + Class.Users.COMPCODE);
                    pp.StartDate = Convert.ToDateTime(dataGridView1.Rows[i].Cells[4].EditedFormattedValue.ToString().Substring(0, 10));
                    pp.EndDate = Convert.ToDateTime(dataGridView1.Rows[i].Cells[5].Value.ToString().Substring(0, 10).Trim());
                    pp.TotalDays = Convert.ToInt64("0"+dataGridView1.Rows[i].Cells[6].Value.ToString().Trim());
                    pp.SalaryDate = Convert.ToDateTime(dataGridView1.Rows[i].Cells[7].EditedFormattedValue.ToString().Substring(0, 10));
                    pp.PayMonth=Convert.ToString(dataGridView1.Rows[i].Cells[8].Value.ToString().Trim());
                    pp.PayPeriod = Convert.ToString(dataGridView1.Rows[i].Cells[9].Value.ToString().Trim());
                    pp.WeeklyHolDay = Convert.ToInt64(dataGridView1.Rows[i].Cells[10].Value.ToString().Trim());
                    pp.WorkingDay = Convert.ToInt64(dataGridView1.Rows[i].Cells[11].Value.ToString().Trim());
                    pp.Notes = Convert.ToString(dataGridView1.Rows[i].Cells[12].EditedFormattedValue.ToString().Trim());
                    sb.Clear(); sb.Append("select asptblfinwekdetid   from  asptblfinwekdet   where   Asptblfinid='" + pp.Asptblfinid + "' and Finyear='" + pp.Finyear + "' and Compcode='" + pp.CompCode + "' and StartDate=date_format('" + pp.StartDate.ToString("yyyy-MM-dd").Substring(0, 10) + "','%Y-%m-%d')  and EndDate=date_format('" + pp.EndDate.ToString("yyyy-MM-dd").Substring(0, 10) + "','%Y-%m-%d')  and TotalDays='" + pp.TotalDays + "' and SalaryDate=date_format('" + pp.SalaryDate.ToString("yyyy-MM-dd").Substring(0, 10) + "','%Y-%m-%d')  and PayMonth='" + pp.PayMonth + "'  and PayPeriod='" + pp.PayPeriod + "' and WeeklyHolDay='" + pp.WeeklyHolDay + "' and WorkingDay='" + pp.WorkingDay + "' and Notes='" + pp.Notes + "' ");
                    DataSet ds1 = Utility.ExecuteSelectQuery(sb.ToString(), "asptblfinwekdet");
                    DataTable dt1 = ds1.Tables["asptblfinwekdet"];
                    if (dt1.Rows.Count != 0)
                    {
                        return;
                    }
                    else if (dt1.Rows.Count != 0 && pp.asptblfinwekdetid == 0 || pp.asptblfinwekdetid == 0)
                    {
                        sb.Clear(); sb.Append("insert into asptblfinwekdet(Asptblfinid,Finyear,CompCode,StartDate,EndDate,TotalDays,SalaryDate,PayMonth,PayPeriod,WeeklyHolDay,WorkingDay,Notes) values('" + pp.Asptblfinid + "','" + pp.Finyear + "','" + pp.CompCode + "',date_format('" + pp.StartDate.ToString("yyyy-MM-dd").Substring(0, 10) + "','%Y-%m-%d'),date_format('" + pp.EndDate.ToString("yyyy-MM-dd").Substring(0, 10) + "','%Y-%m-%d'), '" + pp.TotalDays + "' ,date_format('" + pp.SalaryDate.ToString("yyyy-MM-dd").Substring(0, 10) + "','%Y-%m-%d') ,'" + pp.PayMonth + "','" + pp.PayPeriod + "','" + pp.WeeklyHolDay + "','" + pp.WorkingDay + "','" + pp.Notes + "')");
                        Utility.ExecuteNonQuery(sb.ToString());
                    }
                    else
                    {
                        sb.Clear(); sb.Append("update  asptblfinwekdet   set   Asptblfinid='" + pp.Asptblfinid + "' , Finyear='" + pp.Finyear + "' , Compcode='" + pp.CompCode + "' , StartDate=date_format('" + pp.StartDate.ToString("yyyy-MM-dd").Substring(0, 10) + "','%Y-%m-%d') , EndDate=date_format('" + pp.EndDate.ToString("yyyy-MM-dd").Substring(0, 10) + "','%Y-%m-%d')  , TotalDays='" + pp.TotalDays + "' , SalaryDate=date_format('" + pp.SalaryDate.ToString("yyyy-MM-dd").Substring(0, 10) + "','%Y-%m-%d') ,PayMonth='" + pp.PayMonth + "',PayPeriod='" + pp.PayPeriod + "' , WeeklyHolDay='" + pp.WeeklyHolDay + "' , WorkingDay='" + pp.WorkingDay + "' , Notes='" + pp.Notes + "'  where asptblfinwekdetid='" + pp.asptblfinwekdetid + "' and Finyear='" + pp.Finyear + "' and Compcode='" + pp.CompCode + "' ");
                        Utility.ExecuteNonQuery(sb.ToString());
                    }
                }
            }
           
        }
        private void GridView3(DataGridView dataGridView1)
        {

            int i = 0, j = 1; int cnt = dataGridView1.Rows.Count;
            if (cnt >= 0)
            {
                sb.Clear(); sb.Append("select max(gtfinancialyearid) id    from  gtfinancialyear   where  compcode='" + Class.Users.COMPCODE + "'  and finyear='" + txtfinyear.Text.Trim() + "' ");
                DataSet ds2 = Utility.ExecuteSelectQuery(sb.ToString(), "gtfinancialyear");
                DataTable dt2 = ds2.Tables["gtfinancialyear"];
                maxid = dt2.Rows[0]["id"].ToString();

                for (i = 0; i < cnt; i++)
                {
                    if (dataGridView1.Rows[i].Cells[0].Value == null)
                    {
                        dataGridView1.Rows[i].Cells[0].Value = 0;
                    }
                    else
                    {
                        pp.asptblfinmondetid = Convert.ToInt64("0" + dataGridView1.Rows[i].Cells[0].Value.ToString());
                    }
                    pp.asptblfinmondet1id = Convert.ToInt64("0" + maxid);
                    pp.Asptblfinid = Convert.ToInt64("0" + maxid);
                    pp.Finyear = txtfinyear.Text;
                    pp.CompCode = Convert.ToInt64("0" + Class.Users.COMPCODE);
                    pp.StartDate = Convert.ToDateTime(dataGridView1.Rows[i].Cells[4].EditedFormattedValue.ToString().Substring(0, 10));
                    pp.EndDate = Convert.ToDateTime(dataGridView1.Rows[i].Cells[5].Value.ToString().Substring(0, 10).Trim());
                       //asptblfinmondetid,Asptblfinid,Finyear,CompCode,StartDate,EndDate,PayPeriod,SalaryDate,PayPeriodDays,WeeklyHolDay,Notes
                    pp.PayPeriod = Convert.ToString(dataGridView1.Rows[i].Cells[6].Value.ToString().Trim());
                    pp.SalaryDate = Convert.ToDateTime(dataGridView1.Rows[i].Cells[7].Value.ToString().Trim());
                    pp.PayPeriodDays = Convert.ToInt64(dataGridView1.Rows[i].Cells[8].Value.ToString().Trim());
                    pp.WeeklyHolDay = Convert.ToInt64(dataGridView1.Rows[i].Cells[9].Value.ToString().Trim());
                    pp.Notes = Convert.ToString(dataGridView1.Rows[i].Cells[10].EditedFormattedValue.ToString().Trim());
                    sb.Clear(); sb.Append("select asptblfinmondetid   from  asptblfinmondet   where   Asptblfinid='" + pp.Asptblfinid + "' and Finyear='" + pp.Finyear + "' and Compcode='" + pp.CompCode + "' and StartDate=date_format('" + pp.StartDate.ToString("yyyy-MM-dd").Substring(0, 10) + "','%Y-%m-%d')  and EndDate=date_format('" + pp.EndDate.ToString("yyyy-MM-dd").Substring(0, 10) + "','%Y-%m-%d') and PayPeriod='" + pp.PayPeriod + "'  and SalaryDate=date_format('" + pp.SalaryDate.ToString("yyyy-MM-dd").Substring(0, 10) + "','%Y-%m-%d')   and PayPeriodDays='" + pp.PayPeriodDays + "' and WeeklyHolDay='" + pp.WeeklyHolDay + "'  and Notes='" + pp.Notes + "' ");
                    DataSet ds1 = Utility.ExecuteSelectQuery(sb.ToString(), "asptblfinmondet");
                    DataTable dt1 = ds1.Tables["asptblfinmondet"];
                    if (dt1.Rows.Count != 0)
                    {
                        
                    }
                    else if (dt1.Rows.Count != 0 && pp.asptblfinmondetid == 0 || pp.asptblfinmondetid == 0)
                    {
                        sb.Clear(); sb.Append("insert into asptblfinmondet(Asptblfinid,Finyear,CompCode,StartDate,EndDate,PayPeriod,SalaryDate,PayPeriodDays,WeeklyHolDay,Notes) values('" + pp.Asptblfinid + "','" + pp.Finyear + "','" + pp.CompCode + "',date_format('" + pp.StartDate.ToString("yyyy-MM-dd").Substring(0, 10) + "','%Y-%m-%d'),date_format('" + pp.EndDate.ToString("yyyy-MM-dd").Substring(0, 10) + "','%Y-%m-%d'), '" + pp.PayPeriod + "' ,date_format('" + pp.SalaryDate.ToString("yyyy-MM-dd").Substring(0, 10) + "','%Y-%m-%d') ,'" + pp.PayPeriodDays + "','" + pp.WeeklyHolDay + "','" + pp.Notes + "')");
                        Utility.ExecuteNonQuery(sb.ToString());
                    }
                    else
                    {
                        sb.Clear(); sb.Append("update  asptblfinmondet   set   Asptblfinid='" + pp.Asptblfinid + "' , Finyear='" + pp.Finyear + "' , Compcode='" + pp.CompCode + "' , StartDate=date_format('" + pp.StartDate.ToString("yyyy-MM-dd").Substring(0, 10) + "','%Y-%m-%d')  , EndDate=date_format('" + pp.EndDate.ToString("yyyy-MM-dd").Substring(0, 10) + "','%Y-%m-%d') , PayPeriod='" + pp.PayPeriod + "'  , SalaryDate=date_format('" + pp.SalaryDate.ToString("yyyy-MM-dd").Substring(0, 10) + "','%Y-%m-%d')   , PayPeriodDays='" + pp.PayPeriodDays + "' , WeeklyHolDay='" + pp.WeeklyHolDay + "'  , Notes='" + pp.Notes + "'  where asptblfinmondetid='" + pp.asptblfinmondetid + "' and Finyear='" + pp.Finyear + "' and Compcode='" + pp.CompCode + "' ");
                        Utility.ExecuteNonQuery(sb.ToString());
                    }
                }
            }

        }
        private void empty()
        {
            txtfinyearid.Text = "";
            txtfinyear.Text = "";
            txtfromdate.Text = null; GlobalVariables.New_Flg = false;
            txttodate.Text = "";
            txttotaldays.Text = "";
            butheader.BackColor = Class.Users.BackColors;
            panel3.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
            this.BackColor = Class.Users.BackColors;
            this.Font = Class.Users.FontName;
            dataGridView1.Columns.Clear(); dataGridView2.Columns.Clear(); dataGridView3.Columns.Clear();
            Class.Users.TableName = "asptblfin";
            Class.Users.TableNameGrid = "asptblfindet";
            GlobalVariables.HideCols = new string[] { "Asptblfindetid", "Asptblfinid", "Finyear", "Compcode" };
            GlobalVariables.WidthCols = new Int32[] { 0, 0, 0, 0, 150, 150, 150, 180, 120, 120, 100,150 };
            Class.Users.Query = "select  Asptblfindetid,Asptblfinid,Finyear,Compcode,PeriodStartDate,PeriodEndDate,PeriodCode,PeriodDescription,SalCalDesc,Quarter,TotalDays,Notes from asptblfindet where asptblfindetid<0";
            CommonFunctions.AddGridColumn(dataGridView1, Class.Users.Query, GlobalVariables.HideCols, GlobalVariables.WidthCols, "asptblfindet");

            //Class.Users.Query = ""; Class.Users.TableNameGrid = ""; GlobalVariables.HideCols = null; GlobalVariables.WidthCols = null;
            //Class.Users.TableNameGrid = "asptblfinwekdet";
            GlobalVariables.HideCols = new string[] { "asptblfinwekdetid", "Asptblfinid", "Finyear", "CompCode" };
            GlobalVariables.WidthCols = new Int32[] { 0, 0, 0, 0, 100,100, 100,120, 150, 180,120,120,120 };
            Class.Users.Query = "select  asptblfinwekdetid,Asptblfinid,Finyear,CompCode,startdate,EndDate,TotalDays,SalaryDate,PayMonth,PayPeriod,WeeklyHolDay,WorkingDay,Notes from asptblfinwekdet where asptblfinwekdetid<0";
            CommonFunctions.AddGridColumn(dataGridView2, Class.Users.Query, GlobalVariables.HideCols, GlobalVariables.WidthCols, "asptblfinwekdet");
            //Class.Users.Query = ""; Class.Users.TableNameGrid = ""; GlobalVariables.HideCols = null; GlobalVariables.WidthCols = null;
            //Class.Users.TableNameGrid = "asptblfinmondet";
            GlobalVariables.HideCols = new string[] { "asptblfinmondetid", "Asptblfinid", "Finyear", "CompCode" };
            GlobalVariables.WidthCols = new Int32[] { 0, 0, 0, 0, 110, 110, 150, 130, 200, 150,120 };
            Class.Users.Query = "select  asptblfinmondetid,Asptblfinid,Finyear,CompCode,startdate,EndDate,PayPeriod,SalaryDate,PayPeriodDays,WeeklyHolDay,Notes from asptblfinmondet where asptblfinmondetid<0";
            CommonFunctions.AddGridColumn(dataGridView3, Class.Users.Query, GlobalVariables.HideCols, GlobalVariables.WidthCols, "asptblfinmondet");
            dataGridView1.Controls.Add(dtp); dataGridView1.Controls.Add(dtp01); dataGridView2.Controls.Add(dtp2); dataGridView2.Controls.Add(dtp3); dataGridView2.Controls.Add(dtp4); dataGridView1.Controls.Add(dtp1);
            dataGridView3.Controls.Add(dtpgrid31); dataGridView3.Controls.Add(dtpgrid32); dataGridView3.Controls.Add(dtpgrid33);
            dtp.Visible = false; dtp1.Visible = false;
            dtp.Format = DateTimePickerFormat.Short;

            dtp01.Visible = false; dtp01.Visible = false;
            dtp01.Format = DateTimePickerFormat.Short;
            dtp01.TextChanged += dtp01_TextChanged;
            dtp2.Visible = false;
            dtp2.Format = DateTimePickerFormat.Short;
            dtp2.TextChanged += Dtp2_TextChanged;
            dtp3.Visible = false;
            dtp3.Format = DateTimePickerFormat.Short;
            dtp3.TextChanged += Dtp3_TextChanged;
            dtp4.Visible = false;
            dtp4.Format = DateTimePickerFormat.Short;
            dtp4.TextChanged += Dtp4_TextChanged;

            dtpgrid31.Visible = false;
            dtpgrid31.Format = DateTimePickerFormat.Short;
            dtpgrid31.TextChanged += dtpgrid31_TextChanged;
           
            dtpgrid32.Visible = false;
            dtpgrid32.Format = DateTimePickerFormat.Short;
            dtpgrid32.TextChanged += dtpgrid32_TextChanged;

            dtpgrid33.Visible = false;
            dtpgrid33.Format = DateTimePickerFormat.Short;
            dtpgrid33.TextChanged += dtpgrid33_TextChanged;

            dtp.TextChanged += Dtp_TextChanged;
            dtp1.SelectedIndexChanged += Dtp1_SelectedIndexChanged;
            this.dtp1.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.dtp1.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.dtp1.DropDownHeight = 90;
            listView1.Font = Class.Users.FontName;
          
        }

       

        bool valid = false;
        DateTimePicker dtp = new DateTimePicker();
        DateTimePicker dtp01 = new DateTimePicker();
        DateTimePicker dtp2 = new DateTimePicker();
        DateTimePicker dtp3 = new DateTimePicker();
        DateTimePicker dtp4 = new DateTimePicker();
        DateTimePicker dtpgrid31 = new DateTimePicker();
        DateTimePicker dtpgrid32 = new DateTimePicker();
        DateTimePicker dtpgrid33 = new DateTimePicker();
        ComboBox dtp1 = new ComboBox();
        Rectangle rectangle; 
        string day1 = "";
        string month1 = "";
        string year1 = "";
        int ind = 0,m=0; string griddelrow = "";
        private void Dtp_TextChanged(object sender, EventArgs e)
        {
            try
            {
                m = 0;
                m = ind;
                if (ind <= 11)
                {
                    
                    dataGridView1.CurrentCell.Value = dtp.Value.ToString("dd-MM-yyyy");
                    day1 = dtp.Value.ToString("dddd");
                    mas.checkduplicate2(4, dataGridView1); mas.checkduplicate2(5, dataGridView1);
                    month1 = dtp.Value.ToString("MMMM") + "-" + dtp.Value.ToString("yyyy");
                    year1 = dtp.Value.ToString("yyyy");
                    dataGridView1.Rows[ind].DefaultCellStyle.BackColor = ind % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                    dataGridView1.Rows[ind].Cells[6].Value = "0" + Convert.ToInt32(m + 1);
                    dataGridView1.Rows[ind].Cells[7].Value = year1 + " - 0" + Convert.ToInt32(m + 1);
                    dataGridView1.Rows[ind].Cells[8].Value = "  M" + Convert.ToInt32(m + 1);
                    StartingDate = Convert.ToDateTime(dataGridView1.Rows[ind].Cells[4].Value);
                    EndingDate = Convert.ToDateTime(dataGridView1.Rows[ind].Cells[5].Value);
                    countdays = EndingDate.Subtract(StartingDate);
                    int totaldays = Convert.ToInt32(countdays.Days+1);
                    if (ind < 3)
                    {
                        dataGridView1.Rows[ind].Cells[9].Value = "  Q" + Convert.ToInt32(1);
                    }
                    if (ind >= 3 && ind <= 6)
                    {
                        dataGridView1.Rows[ind].Cells[9].Value = "  Q" + Convert.ToInt32(2);
                    }
                    if (ind >= 6 && ind <= 9)
                    {
                        dataGridView1.Rows[ind].Cells[9].Value = "  Q" + Convert.ToInt32(3);
                    }
                    if (ind >= 9 && ind <= 12)
                    {
                        dataGridView1.Rows[ind].Cells[9].Value = "  Q" + Convert.ToInt32(4);
                    }
                    if (totaldays > 0)
                    {
                        dataGridView1.Rows[ind].Cells[10].Value = totaldays;
                      
                    }
                    
                    totaldays = 0;
                }
                else
                {
                    return;
                }
            }
            catch (Exception ex) { }




        }
        private void dtp01_TextChanged(object sender, EventArgs e)
        {
            try
            {
                m = 0;
                m = ind;

                dataGridView1.CurrentCell.Value = dtp01.Value.ToString("dd-MM-yyyy");

                mas.checkduplicate2(5, dataGridView1);
            }
            catch (Exception ex) { }




        }
        private void Dtp1_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView1.Rows[ind].Cells[11].Value = dtp1.Text.ToString();
            dtp.Visible = false;
            dtp01.Visible = false;
        }
        private void Dtp2_TextChanged(object sender, EventArgs e)
        {
            try
            {
                m = 0;
                m = ind;

                dataGridView2.Rows[ind].Cells[4].Value = dtp2.Value.ToString("dd-MM-yyyy");
                      
                    dataGridView2.Rows[ind].DefaultCellStyle.BackColor = ind % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;


            }
            catch (Exception ex) { }




        }
        private void Dtp3_TextChanged(object sender, EventArgs e)
        {
            try
            {
                m = 0;
                m = ind;
                if (ind <= 11)
                {
                    dataGridView2.Rows[ind].Cells[5].Value = dtp3.Value.ToString("dd-MM-yyyy");
                    dtp3.Visible = false;

                }

            }
            catch (Exception ex) { }




        }
        private void Dtp4_TextChanged(object sender, EventArgs e)
        {
            try
            {
                m = 0;
                m = ind;
                if (ind <= 11)
                {
                    dataGridView2.Rows[ind].Cells[7].Value = dtp4.Value.ToString("dd-MM-yyyy");
                    dtp4.Visible = false;
                }

            }
            catch (Exception ex) { }




        }
        private void dtpgrid31_TextChanged(object sender, EventArgs e)
        {
            try
            {
                m = 0;
                m = ind;
                if (ind <= 11)
                {
                    dataGridView3.Rows[ind].Cells[4].Value = dtpgrid31.Value.ToString("dd-MM-yyyy");
                    dataGridView3.Rows[ind].DefaultCellStyle.BackColor = ind % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;
                    dtpgrid31.Visible = false;
                }

            }
            catch (Exception ex) { }




        }
        private void dtpgrid32_TextChanged(object sender, EventArgs e)
        {
            try
            {
                m = 0;
                m = ind;
                if (ind <= 11)
                {
                    dataGridView3.Rows[ind].Cells[5].Value = dtpgrid32.Value.ToString("dd-MM-yyyy");
                    dtpgrid32.Visible = false;
                }

            }
            catch (Exception ex) { }




        }
        private void dtpgrid33_TextChanged(object sender, EventArgs e)
        {
            try
            {
                m = 0;
                m = ind;
                if (ind <= 11)
                {
                    dataGridView3.Rows[ind].Cells[7].Value = dtpgrid33.Value.ToString("dd-MM-yyyy");
                        dtpgrid33.Visible = false;
                }

            }
            catch (Exception ex) { }




        }
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {

                if (e.ColumnIndex == 4)
                {
                    //case 5: // Column index of needed dateTimePicker cell

                    rectangle = dataGridView1.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true); //  
                    dtp.Size = new Size(rectangle.Width, rectangle.Height); //  
                    dtp.Location = new Point(rectangle.X, rectangle.Y); //  
                    dtp.Visible = true;
                    ind = e.RowIndex;
                }
                if (e.ColumnIndex == 5)
                {
                    rectangle = dataGridView1.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true); //  
                    dtp.Size = new Size(rectangle.Width, rectangle.Height); //  
                    dtp.Location = new Point(rectangle.X, rectangle.Y); //  
                    dtp.Visible = true;
                    ind = e.RowIndex;                    
                }
                if (e.ColumnIndex == 6)
                {
                    dtp.Visible = false; dtp01.Visible = false;
                    if (dataGridView1.Rows[ind].Cells[4].Value != null && dataGridView1.Rows[ind].Cells[5].Value != null)
                    {
                        StartingDate = Convert.ToDateTime(dataGridView1.Rows[ind].Cells[4].Value);
                        EndingDate = Convert.ToDateTime(dataGridView1.Rows[ind].Cells[5].Value);
                        TimeSpan timediff = new TimeSpan();
                        timediff = EndingDate.Subtract(StartingDate);
                        int d1 = timediff.Days + 1;
                        if (d1 < 0)
                        {


                            MessageBox.Show("Date Invalid  From : " + dataGridView1.Rows[ind].Cells[4].Value.ToString() + " To " + dataGridView1.Rows[ind].Cells[5].Value.ToString());
                            dataGridView1.Rows[ind].Cells[5].Value = null;
                        }
                    }
                }
                if (e.ColumnIndex == 11)
                {
                    rectangle = dataGridView1.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true); //  
                    dtp1.Size = new Size(rectangle.Width, rectangle.Height); //  
                    dtp1.Location = new Point(rectangle.X, rectangle.Y); //  
                    dtp1.Visible = true;
                    ind = e.RowIndex;
                }
              
                if (txtfinyearid.Text != "")
                {
                    griddelrow = "";
                    griddelrow = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();

                }

            }
            catch (Exception ex) { }
        }
        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                m = 0;
                if (e.ColumnIndex == 4)
                {
                    //case 5: // Column index of needed dateTimePicker cell

                    rectangle = dataGridView2.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true); //  
                    dtp2.Size = new Size(rectangle.Width, rectangle.Height); //  
                    dtp2.Location = new Point(rectangle.X,rectangle.Y); //  
                    dtp2.Visible = true; dtp3.Visible = false;
                    ind = e.RowIndex;
                }
        
               
                if (e.ColumnIndex == 5)
                {
                    rectangle = dataGridView2.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true); //  
                    dtp3.Size = new Size(rectangle.Width, rectangle.Height); //  
                    dtp3.Location = new Point(rectangle.X, rectangle.Y); //  
                    dtp3.Visible = true; dtp2.Visible = false;
                     ind = e.RowIndex;
                }
                if (e.ColumnIndex == 6)
                {
                    dtp2.Visible = false;
                    dtp3.Visible = false;
                    if (dataGridView2.Rows[ind].Cells[4].Value != null && dataGridView2.Rows[ind].Cells[5].Value != null)
                    {
                        StartingDate = Convert.ToDateTime(dataGridView2.Rows[ind].Cells[4].Value);
                        EndingDate = Convert.ToDateTime(dataGridView2.Rows[ind].Cells[5].Value);
                        TimeSpan timediff = new TimeSpan();
                        timediff = EndingDate.Subtract(StartingDate);
                        int d1 = timediff.Days + 1;
                        if (d1 > 0)
                        {
                            day1 = dtp3.Value.ToString("dddd");
                            month1 = dtp3.Value.ToString("MMMM") + "-" + dtp3.Value.ToString("yyyy");
                            year1 = dtp3.Value.ToString("yyyy");
                            int StartDate = Convert.ToInt32(dataGridView2.Rows[ind].Cells[4].Value.ToString().Substring(0, 2));
                            int EndDate = Convert.ToInt32(dataGridView2.Rows[ind].Cells[5].Value.ToString().Substring(0, 2));

                            int Startmonth = Convert.ToInt32(dataGridView2.Rows[ind].Cells[4].Value.ToString().Substring(3, 2));
                            int Endmonth = Convert.ToInt32(dataGridView2.Rows[ind].Cells[5].Value.ToString().Substring(3, 2));
                            int Startyear = Convert.ToInt32(dataGridView2.Rows[ind].Cells[4].Value.ToString().Substring(6, 4));


                            int totaldays = DateTime.DaysInMonth(Startyear, Startmonth);

                            if (Startmonth == Endmonth)
                            {
                                for (int ii = StartDate; ii <= EndDate; ++ii)
                                {
                                    DateTime dd = new DateTime(EndingDate.Year, EndingDate.Month, ii);
                                    if (dd.DayOfWeek == DayOfWeek.Sunday)
                                    {
                                        m = m + 1;
                                    }
                                }
                            }
                            if (Startmonth != Endmonth)
                            {
                                for (int ii = StartDate; ii <= totaldays; ++ii)
                                {
                                    DateTime dd = new DateTime(StartingDate.Year, StartingDate.Month, ii);
                                    if (dd.DayOfWeek == DayOfWeek.Sunday)
                                    {
                                        m = m + 1;
                                    }
                                }
                                for (int ii = 1; ii <= EndDate; ii++)
                                {
                                    DateTime dd = new DateTime(EndingDate.Year, EndingDate.Month, ii);
                                    if (dd.DayOfWeek == DayOfWeek.Sunday)
                                    {
                                        m = m + 1;
                                    }
                                }
                            }
                            dataGridView2.Rows[ind].Cells[6].Value = d1.ToString();
                            dataGridView2.Rows[ind].Cells[8].Value = month1;
                            dataGridView2.Rows[ind].Cells[9].Value = dataGridView2.Rows[ind].Cells[4].Value.ToString() + "/" + dataGridView2.Rows[ind].Cells[5].Value.ToString();
                            dataGridView2.Rows[ind].Cells[10].Value = m;
                            dataGridView2.Rows[ind].Cells[11].Value = d1 - m;
                        }
                        else
                        {
                            MessageBox.Show("Date Invalid  From : " + dataGridView2.Rows[ind].Cells[4].Value.ToString() +" To " + dataGridView2.Rows[ind].Cells[5].Value.ToString());
                            dataGridView2.Rows[ind].Cells[5].Value = null;
                        }
                    }
                    else
                    {

                    }
                }
                if (e.ColumnIndex == 7)
                {
                    rectangle = dataGridView2.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true); //  
                    dtp4.Size = new Size(rectangle.Width, rectangle.Height); //  
                    dtp4.Location = new Point(rectangle.X, rectangle.Y); //  
                    dtp4.Visible = true;
                    ind = e.RowIndex;
                    
                }
                if (e.ColumnIndex == 12)
                {
                  
                    dtp4.Visible = false;
                   
                }
                if (txtfinyearid.Text != "")
                {
                    griddelrow = "";
                    griddelrow = dataGridView2.Rows[e.RowIndex].Cells[1].Value.ToString();

                }

            }
            catch (Exception ex) { }
        }
        private void dataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                m = 0;
                if (e.ColumnIndex == 4)
                {
                    //case 5: // Column index of needed dateTimePicker cell

                    rectangle = dataGridView3.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true); //  
                    dtpgrid31.Size = new Size(rectangle.Width, rectangle.Height); //  
                    dtpgrid31.Location = new Point(rectangle.X, rectangle.Y); //  
                    dtpgrid31.Visible = true;
                    ind = e.RowIndex;
                }

                if (e.ColumnIndex == 5)
                {
                    rectangle = dataGridView3.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true); //  
                    dtpgrid32.Size = new Size(rectangle.Width, rectangle.Height); //  
                    dtpgrid32.Location = new Point(rectangle.X, rectangle.Y); //  
                    dtpgrid32.Visible = true;
                    ind = e.RowIndex;
                }

                if (e.ColumnIndex == 7)
                {
                    rectangle = dataGridView3.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true); //  
                    dtpgrid33.Size = new Size(rectangle.Width, rectangle.Height); //  
                    dtpgrid33.Location = new Point(rectangle.X, rectangle.Y); //  
                    dtpgrid33.Visible = true;
                    ind = e.RowIndex;

                }
                if (e.ColumnIndex == 6)
                {
                    dtpgrid32.Visible = false; dtpgrid31.Visible = false;
                    if (dataGridView3.Rows[ind].Cells[4].Value != null && dataGridView3.Rows[ind].Cells[5].Value != null)
                    {
                        StartingDate = Convert.ToDateTime(dataGridView3.Rows[ind].Cells[4].Value);
                        EndingDate = Convert.ToDateTime(dataGridView3.Rows[ind].Cells[5].Value);
                        TimeSpan timediff = new TimeSpan();
                        timediff = EndingDate.Subtract(StartingDate);
                        int d1 = timediff.Days + 1;
                        if (d1 > 0)
                        {
                            day1 = dtpgrid31.Value.ToString("dddd");
                            month1 = dtpgrid31.Value.ToString("MMMM") + "-" + dtpgrid31.Value.ToString("yyyy");
                            year1 = dtpgrid31.Value.ToString("yyyy");
                            int StartDate = Convert.ToInt32(dataGridView3.Rows[ind].Cells[4].Value.ToString().Substring(0, 2));
                            int EndDate = Convert.ToInt32(dataGridView3.Rows[ind].Cells[5].Value.ToString().Substring(0, 2));

                            int Startmonth = Convert.ToInt32(dataGridView3.Rows[ind].Cells[4].Value.ToString().Substring(3, 2));
                            int Endmonth = Convert.ToInt32(dataGridView3.Rows[ind].Cells[5].Value.ToString().Substring(3, 2));
                            int Startyear = Convert.ToInt32(dataGridView3.Rows[ind].Cells[4].Value.ToString().Substring(6, 4));


                            int totaldays = DateTime.DaysInMonth(Startyear, Startmonth);

                            if (Startmonth == Endmonth)
                            {
                                for (int ii = StartDate; ii <= EndDate; ++ii)
                                {
                                    DateTime dd = new DateTime(EndingDate.Year, EndingDate.Month, ii);
                                    if (dd.DayOfWeek == DayOfWeek.Sunday)
                                    {
                                        m = m + 1;
                                    }
                                }
                            }
                            if (Startmonth != Endmonth)
                            {
                                for (int ii = StartDate; ii <= totaldays; ++ii)
                                {
                                    DateTime dd = new DateTime(StartingDate.Year, StartingDate.Month, ii);
                                    if (dd.DayOfWeek == DayOfWeek.Sunday)
                                    {
                                        m = m + 1;
                                    }
                                }
                                for (int ii = 1; ii <= EndDate; ii++)
                                {
                                    DateTime dd = new DateTime(EndingDate.Year, EndingDate.Month, ii);
                                    if (dd.DayOfWeek == DayOfWeek.Sunday)
                                    {
                                        m = m + 1;
                                    }
                                }
                            }
                            dataGridView3.Rows[ind].Cells[6].Value = month1.ToString();

                            dataGridView3.Rows[ind].Cells[8].Value = d1 - m;
                            dataGridView3.Rows[ind].Cells[9].Value = m;
                        }
                        else
                        {
                            MessageBox.Show("Date Invalid  From : " + dataGridView3.Rows[ind].Cells[4].Value.ToString() + " To " + dataGridView3.Rows[ind].Cells[5].Value.ToString());
                            dataGridView3.Rows[ind].Cells[5].Value = null;
                        }
                }
                    else { }
                }
                if (e.ColumnIndex == 10)
                {
                    dtpgrid31.Visible = false;
                    dtpgrid32.Visible = false;
                    dtpgrid33.Visible = false;
                }
                if (txtfinyearid.Text != "")
                {
                    griddelrow = "";
                    griddelrow = dataGridView3.Rows[e.RowIndex].Cells[1].Value.ToString();

                }

            }
            catch (Exception ex) { }
        }
        private void dataGridView1_CellValidated(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 4 || e.ColumnIndex == 5 || e.ColumnIndex == 7)
            {
                mas.checkduplicate2(e.ColumnIndex, dataGridView1);
            }


        }
        private void dataGridView2_CellValidated(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 4 || e.ColumnIndex == 5 || e.ColumnIndex == 6)
            {
                mas.checkduplicate2(e.ColumnIndex, dataGridView2);
            }

        }

        private void dataGridView3_CellValidated(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 4 || e.ColumnIndex == 5 || e.ColumnIndex == 8)
            {
                mas.checkduplicate2(e.ColumnIndex, dataGridView3);
            }
        }

  

        private void ListView1_ItemActivate(object sender, EventArgs e)
        {
            try
            {
                if (listView1.Items.Count > 0)
                {

                    txtfinyearid.Text = listView1.SelectedItems[0].SubItems[2].Text;
                    string sel1 = " select a.gtfinancialyearid, a.finyear, a.fromdate,a.TODATE,a.TOTALDAYS,a.CURRENTFINYR,b.compcode from gtfinancialyear a join gtcompmast b on a.compcode=b.gtcompmastid  where a.gtfinancialyearid=" + txtfinyearid.Text;
                    DataSet ds = Utility.ExecuteSelectQuery(sel1, "gtfinancialyear");
                    DataTable dt = ds.Tables["gtfinancialyear"];

                    if (dt.Rows.Count > 0)
                    {
                        GlobalVariables.New_Flg = true;
                        txtfinyearid.Text = Convert.ToString(dt.Rows[0]["gtfinancialyearid"].ToString());
                        txtfinyear.Text = Convert.ToString(dt.Rows[0]["finyear"].ToString());
                        txtfromdate.Text = Convert.ToString(dt.Rows[0]["FROMDATE"].ToString());
                        txttodate.Text = Convert.ToString(dt.Rows[0]["TODATE"].ToString());
                       txttotaldays.Text = Convert.ToString(dt.Rows[0]["TOTALDAYS"].ToString());
                        if (dt.Rows[0]["CURRENTFINYR"].ToString() == "T") { radiocurrent.Checked = true; } else { radioclosed.Checked = true; radiocurrent.Checked = false; }

                       combo_compcode.Text = dt.Rows[0]["compcode"].ToString();
                 
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void Exit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

       

        private void ToolStripMenuItem1_Click(object sender, EventArgs e)
        {
           
        }

        public void News()
        {
            empty(); GridLoad(); 
        }

       

        public void Prints()
        {
            throw new NotImplementedException();
        }

        public void Searchs()
        {
            
        }

        public void Deletes()
        {
            throw new NotImplementedException();
        }

        public void ReadOnlys()
        {
            throw new NotImplementedException();
        }

        public void Imports()
        {
            throw new NotImplementedException();
        }

        public void Pdfs()
        {
            throw new NotImplementedException();
        }

        public void ChangePasswords()
        {
            throw new NotImplementedException();
        }

        public void DownLoads()
        {
            throw new NotImplementedException();
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
            News();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);
            this.Hide();
        }
       
       
      
       
        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }

     
        private void txtsearch_TextChanged(object sender, EventArgs e)
        {
            try
            {


                int item0 = 0; listView1.Items.Clear();
                if (txtsearch.Text.Length > 1)
                {

                    foreach (ListViewItem item in listfilter.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        if (item.SubItems[3].ToString().Contains(txtsearch.Text))
                        {


                            list.Text = item.SubItems[0].Text;
                            list.SubItems.Add(item.SubItems[1].Text);
                            list.SubItems.Add(item.SubItems[2].Text);
                            list.SubItems.Add(item.SubItems[3].Text);
                            list.SubItems.Add(item.SubItems[4].Text);
                            list.SubItems.Add(item.SubItems[5].Text);
                            list.SubItems.Add(item.SubItems[6].Text);
                            list.SubItems.Add(item.SubItems[7].Text);
                            list.SubItems.Add(item.SubItems[8].Text);
                            list.BackColor = item0 % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

                            listView1.Items.Add(list);


                        }
                        item0++;
                    }
                    lbltotal.Text = "Total Count: " + listView1.Items.Count;
                }
                else
                {

                    ListView ll = new ListView();
                    item0 = listfilter.Items.Count;
                    listView1.Items.Clear();
                   
                    foreach (ListViewItem item in listfilter.Items)
                    {
                        ListViewItem list = new ListViewItem();



                        list.Text = item.SubItems[0].Text;
                        list.SubItems.Add(item.SubItems[1].Text);
                        list.SubItems.Add(item.SubItems[2].Text);
                        list.SubItems.Add(item.SubItems[3].Text);
                        list.SubItems.Add(item.SubItems[4].Text);
                        list.SubItems.Add(item.SubItems[5].Text);
                        list.SubItems.Add(item.SubItems[6].Text);
                        list.SubItems.Add(item.SubItems[7].Text);
                        list.SubItems.Add(item.SubItems[8].Text);
                        if (item0 % 2 == 0) { list.BackColor = Color.White; } else { list.BackColor = Color.WhiteSmoke; }
                        listView1.Items.Add(list);
                        item0++;
                    }
                    lbltotal.Text = "Total Count: " + listView1.Items.Count;
                }



            }
            catch (Exception ex)
            {
                MessageBox.Show("---" + ex.ToString());
            }
        }

       
    }
}

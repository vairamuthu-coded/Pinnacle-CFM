using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Tulpep.NotificationWindow;
using System.Drawing;
using System.Windows.Forms;

namespace Pinnacle.Models
{
    class Master
    {
        public  System.Text.StringBuilder sb { get; set; }
        public string panel1 { get; set; }
        public string panel2 { get; set; }

        public void DateTimeChange(TextBox txt)
        {
            if (txt.Text.Length >= 5)
            {
                string[] ss = txt.Text.Split(':');
                if (ss.Length > 1)
                {
                    int s1 = Convert.ToInt32("0" + ss[0].ToString());
                    int s2 = Convert.ToInt32("0" + ss[1].ToString());
                   
                    if (s1 >= 24 || s2 >= 59) { MessageBox.Show("Invalid  ." + txt.Text); }
                }
                else
                {
                    MessageBox.Show("Invalid  ." + txt.Text);
                }
            }
        }
        public Tulpep.NotificationWindow.PopupNotifier pop(string s0,string s1,string s2)
        {
            Tulpep.NotificationWindow.PopupNotifier p = new PopupNotifier();
            p = new PopupNotifier();p.IsRightToLeft = false;
            p.Image = Properties.Resources.info2;
            p.BodyColor = Class.Users.Color1;
            p.HeaderColor = Class.Users.BackColors;
            p.ContentHoverColor = Class.Users.BackColors;
            p.HeaderHeight = 20;p.ShowOptionsButton = true;
            p.TitleColor = Class.Users.BackColors;
            p.TitleFont = new Font("Arial", 15, FontStyle.Bold);
            p.TitleText = s0;          
            p.ContentFont = Class.Users.FontName;
            p.ContentText = " " + s1 + "   " + s2 + "";
            p.Popup();
            return p;
        }
        List<string> _siz = new List<string>();
        List<string> _col = new List<string>();
        List<string> _head = new List<string>();

        public List<string> SizeIndex
        {
            get
            { return _siz; }
          
        }
        public List<string> ColIndex
        {
            get
            { return _col; }

        }
     
        public List<string> GridHeader
        {
            get
            { return _head; }

        }
        public bool checkduplicate1(int index, DataGridView dgrid,string s)
        {
            // dgrid = new DataGridView();
            foreach (DataGridViewRow row in dgrid.Rows)
            {
                
                //if (dgrid.CurrentCell.Value == null)
                //{
                //    continue;
                //}
                if (row.Cells[index].Value != null && row.Cells[index].Value.ToString() == s)
                {
                    MessageBox.Show("Duplicate not allowed"); MessageBox.Show("Duplicate not allowed");
                    s = "";
                    
                    return false;
                }
            }
            return true;
        }
        public void checkduplicate(int index, DataGridView dgrid)
        {
            // dgrid = new DataGridView();
            foreach (DataGridViewRow row in dgrid.Rows)
            {
                if (row.Index == dgrid.CurrentCell.RowIndex)
                {
                    continue;

                }
                if (dgrid.CurrentCell.Value == null)
                {
                    continue;

                }
                if (row.Cells[index].Value != null && row.Cells[index].Value == dgrid.CurrentCell.Value)
                {

                    MessageBox.Show("Duplicate not allowed ");
                    row.Cells[index].Value = null;
                     dgrid.AllowUserToAddRows = false;



                }
                else
                {
                    dgrid.AllowUserToAddRows = true;
                }
            }
        }
        public void checkduplicate2(int index, DataGridView dgrid)
        {
            // dgrid = new DataGridView();
            foreach (DataGridViewRow row in dgrid.Rows)
            {
                if (row.Index == dgrid.CurrentCell.RowIndex)
                {
                    continue;
                    
                }
                if (dgrid.CurrentCell.Value == null)
                {
                    continue;

                }
                if (row.Cells[index].Value != null && row.Cells[index].EditedFormattedValue.ToString() == dgrid.CurrentCell.EditedFormattedValue.ToString())//dgrid.Rows[dgrid.CurrentCell.RowIndex].Cells[dgrid.CurrentCell.ColumnIndex].FormattedValue.ToString()
                {                   
                    MessageBox.Show("Duplicate not allowed ");
                    dgrid.Rows[dgrid.CurrentCell.RowIndex].Cells[index].Value = null;
                    dgrid.Rows[dgrid.CurrentCell.RowIndex].Cells[index].Selected = true;
                 
                }
                else
                {
                    dgrid.AllowUserToAddRows = true;
                }
            }
        }
        public DataTable dept()
        {
            string sel = "select  gtdeptdesgmastID  ,department  from   gtdeptdesgmast";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "gtdeptdesgmast");
            DataTable dt = new DataTable();
            dt = ds.Tables["gtdeptdesgmast"];
            return dt;
        }
        public DataTable dept(Int64 EMPID)
        {
            string sel = "select d.gtdeptdesgmastid as asptbldepid ,d.mnname1 department from hremploymast a join hremploydetails b on a.hremploymastid = b.hremploymastid join gtcompmast c on a.compcode = c.gtcompmastid join gtdeptdesgmast d on d.gtdeptdesgmastid = b.deptname where b.hremploymastid= " + EMPID+";";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "HREMPLOYDETAILS");
            DataTable dt = new DataTable();
            dt = ds.Tables["HREMPLOYDETAILS"];
            return dt;
        }
        public DataTable username(string s)
        {
            DataTable dt;
            if (Class.Users.HUserName == "VAIRAM")
            {

                string sel1 = "SELECT distinct  D.USERNAME   FROM  ASPTBLMACHINEMAS A   JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE    JOIN ASPTBLMACIP C ON C.ASPTBLMACIPID=A.IPADDRESS   AND C.ACTIVE='T'  JOIN  ASPTBLUSERMAS D ON D.USERID=A.USERNAME WHERE B.COMPCODE='" + s + "'  ORDER BY 1";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLMACHINEMAS");
                dt = ds.Tables["ASPTBLMACHINEMAS"];
            }
            else
            {
                string sel1 = "SELECT distinct  D.USERNAME   FROM  ASPTBLMACHINEMAS A   JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE    JOIN ASPTBLMACIP C ON C.ASPTBLMACIPID=A.IPADDRESS   AND C.ACTIVE='T'  JOIN  ASPTBLUSERMAS D ON D.USERID=A.USERNAME WHERE B.COMPCODE='" + s + "' and d.username='" + Class.Users.HUserName + "' ORDER BY 1";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "ASPTBLMACHINEMAS");
                dt = ds.Tables["ASPTBLMACHINEMAS"];
            }
            return dt;
        }
        public DataTable EmpName()
        {
            string selemp = " SELECT 0 AS  asptblempid,'' AS EMPNAME FROM DUAL UNION ALL select DISTINCT a.hremploymastid as asptblempid,a.fname as empname from hremploymast a join hremploydetails b on a.hremploymastid = b.hremploymastid join gtcompmast c on a.compcode = c.gtcompmastid ORDER BY 1; ";
            DataSet dsemp = Utility.ExecuteSelectQuery(selemp, "HREMPLOYDETAILS");
            DataTable dtemp = dsemp.Tables["HREMPLOYDETAILS"];
            return dtemp;
        }
        public DataTable EmpName(string ccode)
        {
            string selemp = "select DISTINCT a.hremploymastid as asptblempid,CONCAT(a.fname ,concat('-',B.MIDCARD) ) as empname from hremploymast a  join hremploydetails b on a.hremploymastid = b.hremploymastid  join gtcompmast c on a.compcode = c.gtcompmastid join gtdeptdesgmast d on d.gtdeptdesgmastid = b.deptname  where  c.compcode='" + ccode + "'  order by 1; ";
            DataSet dsemp = Utility.ExecuteSelectQuery(selemp, "hremploymast");
            DataTable dtemp = dsemp.Tables["hremploymast"];
            return dtemp;
        }
        public DataTable EmployeeName()
        {
            string selemp = "select a.hremploymastid as asptblempid,CONCAT(a.fname ,concat('-',B.MIDCARD) ) as empname from hremploymast a  join hremploydetails b on a.hremploymastid = b.hremploymastid  join gtcompmast c on a.compcode = c.gtcompmastid join gtdeptdesgmast d on d.gtdeptdesgmastid = b.deptname  order by 1; ";
            DataSet dsemp = Utility.ExecuteSelectQuery(selemp, "hremploymast");
            DataTable dtemp = dsemp.Tables["hremploymast"];
            return dtemp;
        }
        //public DataTable VehEmpName(string ccode)
        //{
        //    string selemp = "SELECT  A.ASPTBLVEHMASID1,A.PARTYNAME FROM ASPTBLVEHMAS A    JOIN GTCOMPMAST C ON A.COMPCODE = C.GTCOMPMASTID  WHERE  C.COMPCODE='" + ccode + "'  ORDER BY 1 ";
        //    DataSet dsemp = Utility.ExecuteSelectQuery(selemp, "HREMPLOYMAST");
        //    DataTable dtemp = dsemp.Tables["HREMPLOYMAST"];
        //    return dtemp;
        //}
        public DataTable finyear()
        {
            string sel = "SELECT  A.GTFINANCIALYEARID,A.finyear FROM GTFINANCIALYEAR A   WHERE A.CURRENTFINYR='T';";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "GTFINANCIALYEAR");
            DataTable dt1 = new DataTable();
            dt1 = ds.Tables["GTFINANCIALYEAR"];
            return dt1;
        }
        public DataTable bunkname(string s)
        {
            string sel = "SELECT A.ASPTBLPETMASID,A.BUNKNAME FROM ASPTBLPETMAS A JOIN GTCOMPMAST B ON A.COMPCODE=B.GTCOMPMASTID WHERE A.ACTIVE='T'  AND B.COMPCODE='" + Class.Users.HCompcode + "';";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLPETMAS");
            DataTable dt = new DataTable();
            dt = ds.Tables["ASPTBLPETMAS"];
            return dt;
        }

        public DataTable user(int s)
        {
            string sel = " SELECT A.USERID, A.USERNAME  FROM asptblusermas  A WHERE A.USERID=" + s+";";
            DataSet ds1 = Utility.ExecuteSelectQuery(sel, "asptblusermas");
            DataTable dt1 = ds1.Tables["asptblusermas"];
            return dt1;
        }
        public DataTable phone()
        {
            string sel = " SELECT DISTINCT A.PHONE FROM asptblvisinfo A  WHERE  A.ACTIVE='T'  ORDER BY  1;";
            DataSet ds1 = Utility.ExecuteSelectQuery(sel, "asptblvisinfo");
            DataTable dt1 = ds1.Tables["asptblvisinfo"];
            return dt1;
        }
        public DataTable phone(string time)
        {
            string sel = " SELECT A.PHONE,A.STARTDATE FROM asptblvisinfo A WHERE  A.ACTIVE='T' AND  A. PHONE ='" + time + "' ORDER BY  1;";
            DataSet ds1 = Utility.ExecuteSelectQuery(sel, "asptblvisinfo");
            DataTable dt1 = ds1.Tables["asptblvisinfo"];
            return dt1;
        }
        public DataTable user(Int64 s)
        {
            string sel = " SELECT A.USERID, A.USERNAME, A.IPADDRES,A.CREATATEDON  FROM asptblusermas  A  WHERE A.USERID='" + s + "'; ";
            DataSet ds1 = Utility.ExecuteSelectQuery(sel, "asptblusermas");
            DataTable dt1 = ds1.Tables["asptblusermas"];
            return dt1;
        }
        public DataTable ApplicationName()
        {
            try
            {
                string sel = " SELECT A.asptblparamid ,A.APPNAME ,A.COMPCODE FROM  asptblparam  A     ORDER BY 2; ";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblparam");
                DataTable dt = ds.Tables["asptblparam"];
                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public DataTable comcode()
        {
            try
            {
                string sel = "select distinct a.gtcompmastid,a.compcode,a.compname from  gtcompmast  a WHERE A.ACTIVE='T' Order by 1 ";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "gtcompmast");
                DataTable dt = ds.Tables["gtcompmast"];
                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
            public DataTable locaion()
            {
                try
                {
                    string sel = " select distinct a.asptbllocmasid,a.location from  asptbllocmas  a  where a.active='T' order by 1 ";
                    DataSet ds = Utility.ExecuteSelectQuery(sel, "asptbllocmas");
                    DataTable dt = ds.Tables["asptbllocmas"];
                    return dt;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
            }
       
        public DataTable aspcomcode()
        {
            try
            {
                string sel = " select a.GTCOMPMASTID  , a.compcode   from  GTCOMPMAST   a    where a.active='T';";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "GTCOMPMAST");
                DataTable dt = ds.Tables["GTCOMPMAST"];
                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public DataTable aspcomcode(string comp)
        {
            try
            {
                string sel = " select a.GTCOMPMASTID  , a.compcode,a.compname   from  GTCOMPMAST   a    where a.active='T' and a.compcode='"+comp+"'; ";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "GTCOMPMAST");
                DataTable dt = ds.Tables["GTCOMPMAST"];
                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public DataTable party()
        {
            try
            {
                string sel = "   select distinct a.gtcompmastid,a.compcode, a.compname from  gtcompmast  a where a.ptransaction = 'PARTY'  order by 1 ; ";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "gtcompmast");
                DataTable dt = ds.Tables["gtcompmast"];
                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public DataTable comcode1()
        {
            try
            {

                string sel = "select distinct a.gtcompmastid,  a.compcode from  gtcompmast a  join asptblusermas b on b.compcode = a.gtcompmastid  where b.active='T' order by 2 ";// AND A.COMPCODE='"+Class.Users.HCompcode+"'
                DataSet ds = Utility.ExecuteSelectQuery(sel, "gtcompmast");
                DataTable dt = ds.Tables["gtcompmast"];
                return dt;

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public DataTable findcomcode(string s)
        {

            string sel = "select distinct a.gtcompmastid,a.compcode,a.compname from  gtcompmast  a  where a.compcode='" + s + "';";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "GTCOMPMAST");
            DataTable dt = ds.Tables["GTCOMPMAST"];
            return dt;
        }
        public DataTable findcomcode1(string s)
        {

            string sel = "select distinct a.compcode from  gtcompmast  a where a.compcode='" + s + "'  UNION ALL         select 'TOKEN CANEL' AS compcode from  DUAL            UNION ALL         select 'FULL' AS compcode from  DUAL            UNION ALL         select 'DIESEL' AS compcode from  DUAL            UNION ALL         select 'PETROL' AS compcode from  DUAL        UNION ALL      select '2T OIL' AS compcode from  DUAL   UNION ALL     select 'ALL' AS compcode from  DUAL ;";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "GTCOMPMAST");
            DataTable dt = ds.Tables["GTCOMPMAST"];
            return dt;
        }
        public DataTable findcomcode(string s,string ss)
        {
            string sel = "select distinct b.userid,b.username,a.gtcompmastid,a.compcode,a.compname  from  gtcompmast a   join asptblusermas b on a.gtcompmastid = b.compcode    where a.compcode='" + s + "'   and b.username='" + ss + "'; ";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "gtcompmast");
            DataTable dt = ds.Tables["gtcompmast"];
            return dt;
        }
      
        public DataTable findcomcode(string s, string ss, string sss)
        {
            string sel = "select distinct b.userid,b.username,a.gtcompmastid,a.compcode,a.compname  from  gtcompmast a join asptblusermas b on a.gtcompmastid = b.compcode    where a.COMPCODE='" + s + "'   and b.username='" + ss + "' AND  C.finyr='" + sss+"';";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "gtcompmast");
            DataTable dt = ds.Tables["gtcompmast"];
            return dt;
        }

        public DataTable finduser(Int64 s, Int64 ss)
        {
            
            string sel = "  select  distinct a.userid, a.username,b.compcode from asptblusermas a join gtcompmast b on a.compcode=b.gtcompmastid  where b.gtcompmastid='" + s + "' and c.userid='" + ss + "';";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "asptblusermas");
            DataTable dt = ds.Tables["asptblusermas"];
            return dt;
        }

        public DataTable comcode2(Int64 s)
        {
            string sel = "  select distinct a.gtcompmastid,a.compcode, a.compname,b.userid, b.username from  gtcompmast  a join asptblusermas b on  a.gtcompmastid=b.compcode   join asptblnavigation c on c.compcode=a.gtcompmastid and c.username=b.userid where  a.gtcompmastid=" + s + " order by 5;";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "gtcompmast");
            DataTable dt = ds.Tables["gtcompmast"];
            return dt;
        }
        public DataTable comcode1(Int64 s)
        {
            string sel = "select distinct a.gtcompmastid,a.compcode, a.compname,b.userid, b.username from  gtcompmast  a join asptblusermas b on  a.gtcompmastid=b.compcode  where  a.gtcompmastid=" + s+";";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "gtcompmast");
            DataTable dt = ds.Tables["gtcompmast"];
            return dt;
        }
        public DataTable compcodefind(string s)
        {
            string sel = " SELECT DISTINCT A.GTCOMPMASTID,A.COMPCODE FROM  GTCOMPMAST  A JOIN ASPTBLUSERMAS B ON  A.GTCOMPMASTID=B.COMPCODE   JOIN ASPTBLPETMAS C ON C.COMPCODE=A.GTCOMPMASTID AND C.ACTIVE='T'   where  a.compcode='" + s+"';";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "gtcompmast");
            DataTable dt = ds.Tables["gtcompmast"];
            return dt;
        }
        public DataTable bunkfind(string s)
        {
            string sel = " SELECT DISTINCT C.ASPTBLPETMASID,C.BUNKNAME FROM  GTCOMPMAST  A    JOIN ASPTBLPETMAS C ON C.COMPCODE=A.GTCOMPMASTID AND C.ACTIVE='T'   where  a.compcode='" + s + "';";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "gtcompmast");
            DataTable dt = ds.Tables["gtcompmast"];
            return dt;
        }
    }
    public abstract class userdetails
    {

        public abstract void deleterecord(int p1);
        public abstract DataTable griddisplay();
        public abstract DataTable bind();
        public abstract DataTable comcode();
        public abstract DataTable comcode(string s);
        public abstract DataTable company();
        public abstract DataTable party();
        public abstract DataTable party(string s);
        public abstract DataTable gridselect(string s);


        public abstract string autoid();

    }

    public class ipaddress : userdetails
    {
        public ipaddress()
        {
        }

        public override void deleterecord(int p1)
        {

            string del = "delete from tblfinmas  where tblfinmasid=" + p1;
            Utility.ExecuteNonQuery(del);
        }
        public override DataTable griddisplay()
        {
            string sel = "select  a.tblfinmasid,a.finyear,a.fromdate, a.todate,a.currentyear,a.totaldays  from tblfinmas a order by 2  ";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "tblfinmas");
            DataTable dt = new DataTable();
            dt = ds.Tables["tblfinmas"];
            return dt;

        }
        public override DataTable comcode()
        {
            throw new NotImplementedException();
        }
        public override DataTable comcode(string s)
        {

            string sel = "SELECT A.gtcompmastID,A.COMPNAME FROM gtcompmast A where A.COMPCODE='" + s + "'";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "gtcompmast");
            DataTable dt = new DataTable();
            dt = ds.Tables["gtcompmast"];
            return dt;
        }
        public override DataTable company()
        {
            string sel = "SELECT DISTINCT A.gtcompmastID, A.COMPCODE FROM gtcompmast A where A.ACTIVE='T' ";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "gtcompmast");
            DataTable dt = new DataTable();
            dt = ds.Tables["gtcompmast"];
            return dt;
        }
        public override DataTable party()
        {
            string sel = "SELECT DISTINCT A.ASPTBLPARID, A.PARTYCODE FROM ASPTBLPAR A  ";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLPAR");
            DataTable dt = new DataTable();
            dt = ds.Tables["ASPTBLPAR"];
            return dt;
        }

        public override DataTable party(string p1)
        {
            string sel = "SELECT  A.ASPTBLPARID, A.PARTY FROM ASPTBLPAR A   WHERE A.PARTYCODE='" + p1 + "'";
            DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLPAR");
            DataTable dt = new DataTable();
            dt = ds.Tables["ASPTBLPAR"];
            return dt;
        }
        public override DataTable gridselect(string s)
        {
            string sel1 = "SELECT  A.ASPTBLPARID,  A.SUPPCODE, A.PARTYCODE, A.PARTYNAME, A.ACTIVE, A.USERNAME, A.IPADDRESS,A.CREATEDBY, A.CREATEDON,  A.MODIFIEDON,  A.PARTYCODE1, A.PARTYSNO  FROM ASPTBLPAR A  WHERE A.ASPTBLPARID='" + s + "'";
            DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "ASPTBLPAR");
            DataTable dt = ds1.Tables["ASPTBLPAR"];
            return dt;
        }
        public override DataTable bind()
        {
            throw new System.NotImplementedException();
        }
        public override string autoid()
        {
            throw new NotImplementedException();
        }
    }
}

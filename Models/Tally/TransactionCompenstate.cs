using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pinnacle.Models.Tally
{
    class TransactionCompenstate
    {
        public string projetid { get; set; }
        public string tblname { get; set; }
        public string tblnamegrid { get; set; }
        public double seqno { get; set; }
        public string des { get; set; }
        public string prefix { get; set; }
        public string row { get; set; }
        public string col { get; set; }
        public string file { get; set; }
        public string query { get; set; }
        public string update { get; set; }
        public DataGridView GridView { get; set; }
        public DataSet ds; public DataTable dt;
        public void dropsequence(string s, string ss)
        {
            tblname = ss; projetid = s;
            query = "drop sequence " + projetid + "." + tblname + "SEQ";
            Utility.ExecuteNonQuery(query);
        }
        public void DropTrigger(string s, string ss)
        {
            tblname = ss; projetid = s;
            query = "drop trigger " + projetid + "." + tblname + "TRI";
            Utility.ExecuteNonQuery(query);
        }
        public void CreateSequence(string s, string ss, double sss)
        {
            tblname = ss; projetid = s; seqno = sss;
            query = "CREATE SEQUENCE  " + projetid + "." + tblname + "SEQ" + "  START WITH " + seqno + "  MAXVALUE 9999999999999  MINVALUE 0  NOCYCLE NOCACHE  NOORDER";
            Utility.ExecuteNonQuery(query);
        }
        public DataTable Maximumid(string PRO, string tblname, string tblname1, string TBL1, string TBL2)
        {           
            query = "SELECT MAX(A."+ tblname + "ID)+1 as " + tblname + "ID FROM  " + PRO + "." + tblname + " A JOIN " + tblname1 + " B ON A." + TBL2 + "=B." + TBL1 + " JOIN GTCOMPMAST C ON C.GTCOMPMASTID = B.COMPCODE WHERE C.COMPCODE ='" + Class.Users.HCompcode+"'";
            DataSet ds = Utility.ExecuteSelectQuery(query, tblname);
            DataTable dt = ds.Tables[0];
            return dt;
        }
        public void CreateTrigger(string s, string ss)
        {
            tblname = ss; projetid = s;
            query = "CREATE TRIGGER   " + s + "TRI" + " BEFORE INSERT ON  " + ss + "." + s + " REFERENCING NEW AS NEW OLD AS OLD FOR EACH ROW DECLARE  " + s + "ID INTEGER;BEGIN " + s + "ID:= 0;   SELECT " + ss + "." + s + "SEQ.NEXTVAL INTO " + s + "ID FROM DUAL;   :NEW." + s + "ID:= " + s + "ID;END " + s + "TRI;";
            Utility.ExecuteNonQuery(query);
        }
        public DataTable SlugData(string proid, string tblgrid, string des, string fi, string pre)
        {
            this.projetid = proid; this.tblnamegrid = tblgrid; this.des = des; this.file = fi; this.prefix = pre; 
            query = "SELECT A.AUTOGENERATEID,A.TX_VIEW_ID,A.DESCRIPTION,A.FIELDNAME,A.PREFIX,A.LASTNO,A.ACTIVESEQUENCE,A.USERID,A.PROJECTID,A.DUPLICATECTRL,A.ZEROPADDING FROM " + projetid + ".AUTOGENERATE A  WHERE A.TX_VIEW_ID='" + tblnamegrid + "'  AND A.DESCRIPTION='" + des + "' AND A.FIELDNAME='" + file + "'  AND A.PREFIX='" + prefix + "'   AND A.PROJECTID='" + projetid + "'";
            ds = Utility.ExecuteSelectQuery(query, "AUTOGENERATE");
            dt = ds.Tables["AUTOGENERATE"]; return dt;
        }
        public DataTable TableMax(string s, string ss)
        {
            this.projetid = ss; this.tblname = s;
            query = "select TO_CHAR(MAX(A." + tblname + "ID)+1) LASTNO  from  " + projetid + "." + tblname + " A";
            ds = Utility.ExecuteSelectQuery(query, tblname);
            dt = ds.Tables[tblname];
            return dt;
        }
        public DataTable SlugData1(string s, string ss)
        {
            this.projetid = ss; this.tblnamegrid = s;
            query = "SELECT  A.ZEROPADDING,A.LASTNO FROM " + projetid + ".AUTOGENERATE A  WHERE A.TX_VIEW_ID='" + tblnamegrid + "'    AND A.PROJECTID='" + projetid + "'";
            ds = Utility.ExecuteSelectQuery(query, "AUTOGENERATE");
            dt = ds.Tables["AUTOGENERATE"];
            return dt;
        }
        public DataTable FindSlug(string s)
        {
            this.tblname = s;
            query = "SELECT COUNT(A.SEQUENCE_NAME) CNT FROM USER_SEQUENCES A WHERE A.SEQUENCE_NAME='" + tblname + "SEQ'";
            ds = Utility.ExecuteSelectQuery(query, s + "SEQ");
            dt = ds.Tables[s + "SEQ"];
            return dt;
        }
        public DataTable FindTrigger(string s, string ss)
        {
            this.tblname = s; this.projetid = ss;
            query = "SELECT COUNT(*) CNT FROM USER_TRIGGERS a WHERE A.TRIGGER_NAME='" + tblname + "TRI' AND A.TABLE_OWNER='" + projetid + "'";
            ds = Utility.ExecuteSelectQuery(query, tblname + "TRI");
            dt = ds.Tables[tblname + "TRI"];
            return dt;
        }
        public DataTable GridSelect(string s,string row,string col)
        {
            this.tblname = s; this.row = row;this.col = col;
            query = "select " + tblname + "ID  from  " + tblname + " WHERE " + row + "='" + col + "'";
            ds = Utility.ExecuteSelectQuery(query, tblname);
            dt = ds.Tables[tblname];
            return dt;
        }
        public DataTable FindDuplicate(string s, string ss, string up)
        {
            this.tblname = s; this.projetid = ss; this.update = up;
            query = "select distinct " + tblname + "ID" + " from " + projetid + "." + tblname + update;
            ds = Utility.ExecuteSelectQuery(query, tblname);
            dt = ds.Tables[tblname];
            return dt;
        }
        public void GridRowRemove(DataGridView grid)
        {
           
            do
            {
                for (int i = 0; i < grid.Rows.Count; i++) { try { grid.Rows.RemoveAt(i); } catch (Exception) { } }
            }
            while (grid.Rows.Count >= 1);
        }
    }
}

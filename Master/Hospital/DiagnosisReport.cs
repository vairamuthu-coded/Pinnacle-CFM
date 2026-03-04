using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pinnacle.Master.Hospital
{
    public partial class DiagnosisReport : Form
    {
        public DiagnosisReport()
        {
            InitializeComponent();
            label7.ForeColor = SystemColors.Highlight;
            butfooter.BackColor = SystemColors.Highlight;
        }
        byte[] stdbytes; byte[] stdbytes1;
       // Report.Hospital.TestReport rd1 = new Report.Hospital.TestReport();
        Report.Hospital.HospitalReport rd = new Report.Hospital.HospitalReport();
        //Report.Hospital.DiagnosisReport rd = new Report.Hospital.DiagnosisReport();
        CrystalDecisions.CrystalReports.Engine.ReportDocument reportdocument = new CrystalDecisions.CrystalReports.Engine.ReportDocument();

        private void DiagnosisReport_Load(object sender, EventArgs e)
        {
            try                 
            {
                crystalReportViewer1.ReportSource = null;
                crystalReportViewer1.Refresh(); Class.Users.dt = null;
                int cnt = 0; int i = 0;
                string sel1 = "SELECT a.asptbldiagnosismasid,c.compcode,c.compname,d.address, d.patientname,d.gender,a.diagonisis, a.symptoms,b.medicine,date_format(a.diagnoseddate,'%d-%m-%Y %h:%m:%s') AS diagnoseddate,d.tokenno,d.asptblregistermasid ,d.patientphoto,E.sign,''labtest,'' labtestitem,E.doctorname,E.asptbldocmasid FROM asptbldiagnosismas a join asptblmedmas b on b.asptbldiagnosismasid=a.asptbldiagnosismasid and b.asptblregistermasid=a.asptblregistermasid join gtcompmast c on c.gtcompmastid=a.compcode    join asptblregistermas d on d.asptblregistermasid=A.asptblregistermasid  and d.asptblregistermasid=b.asptblregistermasid   join asptbldocmas e where e.active='T' AND a.asptbldiagnosismasid= '" + Class.Users.Paramid + "'   ";//and e.asptbldocmasid='" + Class.Users.DoctorName + "' and d.asptblregistermasid='" + Class.Users.PatientName + "'
                DataSet ds1 = Utility.ExecuteSelectQuery(sel1, "asptbldiagnosismas");
                DataTable dt1 = ds1.Tables["asptbldiagnosismas"];
                if (dt1.Rows.Count > 0)
                {

                    Class.Users.dt = dt1.Clone();
                    foreach (DataRow porow in dt1.Rows)
                    {
                        Class.Users.dt.Rows.Add();
                        Class.Users.dt.Rows[cnt]["compcode"] = porow.ItemArray[1].ToString();
                        Class.Users.dt.Rows[cnt]["compname"] = porow.ItemArray[2].ToString();
                        Class.Users.dt.Rows[cnt]["address"] = porow.ItemArray[3].ToString();
                        Class.Users.dt.Rows[cnt]["patientname"] = porow.ItemArray[4].ToString();
                        Class.Users.dt.Rows[cnt]["gender"] = porow.ItemArray[5].ToString();
                        Class.Users.dt.Rows[cnt]["diagonisis"] = porow.ItemArray[6].ToString();
                        Class.Users.dt.Rows[cnt]["symptoms"] = porow.ItemArray[7].ToString();
                        Class.Users.dt.Rows[cnt]["medicine"] = porow.ItemArray[8].ToString();
                        Class.Users.dt.Rows[cnt]["diagnoseddate"] = porow.ItemArray[9].ToString();
                        Class.Users.dt.Rows[cnt]["tokenno"] = porow.ItemArray[10].ToString();
                        Class.Users.dt.Rows[cnt]["asptblregistermasid"] = porow.ItemArray[11].ToString();
                        Class.Users.dt.Rows[cnt]["doctorname"] = porow.ItemArray[16].ToString();
                        Class.Users.dt.Rows[cnt]["asptbldocmasid"] = porow.ItemArray[17].ToString();
                        if (dt1.Rows[0]["patientphoto"].ToString() != "")
                        {
                            stdbytes = (byte[])dt1.Rows[0]["patientphoto"];
                            Class.Users.dt.Rows[cnt]["patientphoto"] = stdbytes;
                        }
                        if (dt1.Rows[0]["sign"].ToString() != "")
                        {
                            stdbytes1 = (byte[])dt1.Rows[0]["sign"];
                            Class.Users.dt.Rows[cnt]["sign"] = stdbytes1;
                        }

                    }

                    string sel2 = "SELECT g.labtest,h.labtestitem  FROM asptbldiagnosismas a  join asptblpatientmas b on a.asptblpatientmasid=b.asptblpatientmasid join gtcompmast c on c.gtcompmastid=a.compcode   join asptblregistermas d on d.asptblpatientmasid=b.asptblpatientmasid   join asptbldocmas e on  e.active='T' join asptbldiatest f on f.asptbldiagnosismasid=a.asptbldiagnosismasid  and a.asptblpatientmasid=b.asptblpatientmasid  join asptbllabtestmas g on g.asptbllabtestmasid=f.asptbllabtestmasid    join asptbllabtestitemmas h on h.labtest=f.asptbllabtestmasid and h.asptbllabtestitemmasid=f.asptbllabtestitemmasid where a.asptbldiagnosismasid= '" + Class.Users.Paramid + "'  ";//and e.asptbldocmasid='" + Class.Users.DoctorName + "' and d.asptblregistermasid='" + Class.Users.PatientName + "'
                    DataSet ds2 = Utility.ExecuteSelectQuery(sel2, "asptbldiagnosismas");
                    DataTable dt2 = ds2.Tables["asptbldiagnosismas"];
                    if (dt2 != null)
                    {
                      
                        foreach (DataRow porow1 in dt2.Rows)
                        {
                            if (cnt >= 1)
                            {
                                Class.Users.dt.Rows.Add();
                            }
                            Class.Users.dt.Rows[cnt]["labtest"] = porow1.ItemArray[0].ToString();
                            Class.Users.dt.Rows[cnt]["labtestitem"] = porow1.ItemArray[1].ToString();
                            cnt++;
                        }
                    }



                    rd.Database.Tables["DataTable1"].SetDataSource(Class.Users.dt);
                    crystalReportViewer1.ReportSource = rd;
                    crystalReportViewer1.Refresh();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
    }
}

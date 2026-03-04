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
    public partial class ListViewMaster : Form,ToolStripAccess
    {
        private static ListViewMaster _instance;
        Models.Master mas = new Models.Master();
      
        Models.UserRights sm = new Models.UserRights();
        Int64 std, std1 = 0; ListView listfilter = new ListView();
        byte[] stdbytes; byte[] votebytes; OpenFileDialog open = new OpenFileDialog();
        int i = 0;
        public static ListViewMaster Instance
        {
            get
            {
                if (_instance == null)
                    _instance = new ListViewMaster();
                GlobalVariables.CurrentForm = _instance;
                return _instance;
            }
        }
        public ListViewMaster()
        {
            InitializeComponent();
            Class.Users.IPADDRESS = GenFun.GetLocalIPAddress();
            Class.Users.CREATED = Convert.ToDateTime(System.DateTime.Now.ToString("dd-MMM-yyyy") + " " + System.DateTime.Now.ToLongTimeString());
            Class.Users.SysDate = Convert.ToString(System.DateTime.Now.ToString("dd-MM-yyyy"));
            Class.Users.SysTime = Convert.ToString(DateTime.Now.ToLongTimeString().ToString());
            GlobalVariables.HeaderName.Text = Class.Users.ScreenName;
            GlobalVariables.CurrentForm = this;
            butheader.Text = Class.Users.ScreenName;
            panel1.BackColor = Class.Users.BackColors;
            butheader.BackColor = Class.Users.BackColors;
            panel2.BackColor = Class.Users.BackColors;
        }

        public void News()
        {
           
            butheader.Text = Class.Users.ScreenName;
            panel1.BackColor = Class.Users.BackColors;
            butheader.BackColor = Class.Users.BackColors;           
            panel2.BackColor = Class.Users.BackColors;
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
            
        }

        public void ChangeSkins()
        {
            
        }

        public void Logins()
        {
           
        }

        public void GlobalSearchs()
        {
            
        }

        public void TreeButtons()
        {
            
        }

        public void Exit()
        {
            GlobalVariables.MdiPanel.Show();
            this.Hide();
            News();
            GlobalVariables.HeaderName.Text = "";
            GlobalVariables.TabCtrl.TabPages.RemoveAt(GlobalVariables.TabCtrl.SelectedIndex);
        }

        private void txtsearch_TextChanged(object sender, EventArgs e)
        {
            try
            {
                int item0 = 0;i = 0;
                if (txtsearch.Text.Length > 0)
                {
                    listView1.Items.Clear();
                    foreach (ListViewItem item in listfilter.Items)
                    {
                        ListViewItem list = new ListViewItem();
                        if (listfilter.Items[item0].SubItems[2].ToString().Contains(txtsearch.Text))
                        {
                            list.Text = listfilter.Items[item0].SubItems[0].Text;
                            list.SubItems.Add(listfilter.Items[item0].SubItems[1].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[2].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[3].Text);
                            list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

                            listView1.Items.Add(list);
                        }
                        item0++;
                    }
                }
                else
                {
                    try
                    {
                        listView1.Items.Clear(); item0 = 0;i = 0;
                        foreach (ListViewItem item in listfilter.Items)
                        {
                            ListViewItem list = new ListViewItem();
                            list.Text = listfilter.Items[item0].SubItems[0].Text;
                            list.SubItems.Add(listfilter.Items[item0].SubItems[1].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[2].Text);
                            list.SubItems.Add(listfilter.Items[item0].SubItems[3].Text);

                            list.BackColor = i % 2 == 0 ? Class.Users.Color1 : Class.Users.Color2;

                            listView1.Items.Add(list);
                            item0++;
                        }
                    }
                    catch (Exception ex)
                    {

                    }
                }
            }
            catch (Exception ex)
            {
                // MessageBox.Show("---" + ex.ToString());
            }
        }

        private void listView1_ItemActivate(object sender, EventArgs e)
        {
            try
            {
                if (listView1.Items.Count > 0)
                {

                    Class.Users.Paramid = Convert.ToInt64(listView1.SelectedItems[0].SubItems[2].Text);
                    this.Hide();
                 //   GlobalVariables.CurrentForm = "{PSSRobot.Master.Hospital.DiagnosisMaster, Text: DiagnosisMaster}";
                    
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
           
        }

        private void ListViewMaster_Load(object sender, EventArgs e)
        {
            GridLoad();
        }

        public void GridLoad()
        {
            try
            {
                listView1.Items.Clear(); listfilter.Items.Clear();
                string sel1 = "SELECT distinct a.asptbldiagnosismasid, c.patientname,c.gender,d.doctorname,b.TOKENNO,a.symptoms, a.diagonisis,a.medicine,c.asptblregistermasid FROM asptbldiagnosismas a   join asptblpatientmas b on a.asptblpatientmasid=b.asptblpatientmasid  join asptblregistermas c on c.asptblregistermasid=a.asptblregistermasid   and c.asptblregistermasid=b.patientname   join asptbldocmas d on d.asptbldocmasid=a.doctorname where a.modifiedon='" + System.DateTime.Now.ToString("yyyy-MM-dd") + "'  order by 1";
                DataSet ds = Utility.ExecuteSelectQuery(sel1, "asptbldiagnosismas");
                DataTable dt = ds.Tables["asptbldiagnosismas"];
                if (dt.Rows.Count > 0)
                {
                    int i = 1;
                    foreach (DataRow myRow in dt.Rows)
                    {
                        ListViewItem list = new ListViewItem();
                        list.SubItems.Add(i.ToString());
                        list.SubItems.Add(myRow["asptbldiagnosismasid"].ToString());
                        list.SubItems.Add(myRow["patientname"].ToString());
                        list.SubItems.Add(myRow["gender"].ToString());
                        list.SubItems.Add(myRow["tokenno"].ToString());
                        listfilter.Items.Add((ListViewItem)list.Clone());
                        if (i % 2 == 0)
                        {
                            list.BackColor = Color.White;
                        }
                        else
                        {
                            list.BackColor = Color.WhiteSmoke;
                        }
                        i++;
                        listView1.Items.Add(list);
                    }
                    lbltotal.Text = "Total Count    :" + listView1.Items.Count;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void Searchs(int EditID)
        {
            throw new NotImplementedException();
        }
    }
}

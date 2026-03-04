using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Pinnacle.Transactions
{
    public partial class PopUpWindow : Form
    {
        string txtcutid, txtcutid1, txtcompcode, txtpono, txtdocid = "";
        public PopUpWindow()
        {
            InitializeComponent();
           
        }
        Models.Master mas = new Models.Master();
        private void PopUpWindow_Load(object sender, EventArgs e)
        {
  
          
            txtcutid = Class.Users.GridID;
            txtcutid = Class.Users.NonGridID;
            txtcompcode = Class.Users.CompCode1;
            txtpono = Class.Users.PoNo;
            txtdocid = Class.Users.DocID;
            string sel4 = "select h.Asptblsizmasid,f.Fabric,g.colorname,h.SizeName,e.Markerno,a.Orderqty from asptblpur a join gtcompmast b on a.compcode=b.gtcompmastid  join asptblpurdet c on c.asptblpurid=a.asptblpurid  and c.compcode=a.compcode and c.compcode=b.gtcompmastid  join asptbllay d on d.pono=a.pono   join asptbllaydet e on e.asptbllayid=d.asptbllayid and e.asptbllayid1=d.asptbllayid1  and e.colorname=c.colorname and e.sizename=c.sizename  join asptblfabmas f on f.asptblfabmasid=e.FABRIC join asptblcolmas g on g.asptblcolmasid=e.colorname join asptblsizmas h on h.ASPTBLSIZMASID=e.sizename  where  b.compcode='" + txtcompcode + "' and a.pono='" + txtpono + "' and d.docid='" + txtdocid + "' order by 1";
            DataSet ds4 = Utility.ExecuteSelectQuery(sel4, "asptblcut");
            DataTable dt4 = ds4.Tables["asptblcut"];
            Class.Users.TableNameGrid = "asptblcut";
            if (dt4.Rows.Count > 0)
            {
                int cnt = 0;
               

                foreach (DataColumn item1 in dt4.Columns)
                {
                    mas.ColIndex.Add(item1.ColumnName);
                }
                int i = 0;
                foreach (DataRow item2 in dt4.Rows)
                {
                    mas.SizeIndex.Add(item2[i].ToString());
                    i++;
                }

                if (mas.SizeIndex.Count > 0)
                {
                    CommonFunctions.AddColumnDynamic(dataGridView1,mas.ColIndex.ToArray(), mas.SizeIndex.ToArray(), sel4, GlobalVariables.HideCols,GlobalVariables.WidthCols);
                }
               
                mas.ColIndex.Clear(); mas.SizeIndex.Clear(); mas.GridHeader.Clear(); mas.ColIndex.Clear();
            }

            this.Font = Class.Users.FontName;
            dataGridView1.Font= Class.Users.FontName;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Class.Users.BackColors;

            //for (int i = 0; i < dt4.Rows.Count; i++)
            //{
            //    dataGridView1.Rows.Add();
            //    dataGridView1.Rows[i].Cells[0].Value = dt4.Rows[i]["asptblcutdetid"].ToString();
            //    dataGridView1.Rows[i].Cells[1].Value = dt4.Rows[i]["asptblcutid"].ToString();
            //    dataGridView1.Rows[i].Cells[2].Value = dt4.Rows[i]["asptblcutid1"].ToString();
            //    dataGridView1.Rows[i].Cells[3].Value = dt4.Rows[i]["compcode"].ToString();
            //    dataGridView1.Rows[i].Cells[4].Value = dt4.Rows[i]["layno1"].ToString();
            //    dataGridView1.Rows[i].Cells[5].Value = dt4.Rows[i]["FABRIC"].ToString();
            //    dataGridView1.Rows[i].Cells[6].Value = dt4.Rows[i]["colorname"].ToString();
            //    dataGridView1.Rows[i].Cells[7].Value = dt4.Rows[i]["SIZENAME"].ToString();
            //    dataGridView1.Rows[i].Cells[8].Value = dt4.Rows[i]["markerno"].ToString();
            //    dataGridView1.Rows[i].Cells[9].Value = dt4.Rows[i]["cutqty"].ToString();
            //    dataGridView1.Rows[i].Cells[10].Value = dt4.Rows[i]["Notes"].ToString();
            //}
               //CommonFunctions.SetRowNumber(dataGridView1);
            //}
            //else
            //{
            //    do
            //    {
            //        for (int i = 0; i < dataGridView1.Rows.Count; i++) { try { dataGridView1.Rows.RemoveAt(i); } catch (Exception) { } }
            //    }
            //    while (dataGridView1.Rows.Count > 1);



        }

        
    }
}

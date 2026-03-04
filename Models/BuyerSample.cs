using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Pinnacle.Models
{
  public  class BuyerSample
    {


        public Int64 ASPTBLBUYSAMID { get; set; }
        public string buyercode { get; set; }
        public string styref { get; set; }
        public string styno { get; set; }
        public string fabric { get; set; }
        public string content { get; set; }
        public string counts { get; set; }
        public string gg { get; set; }
        public string gsm { get; set; }
        public string Color { get; set; }
        public byte[] bytes { get; set; }
     
        public string remarks { get; set; }
        public BuyerSample()
        {

        }

        //public BuyerSample(string buyercode, string styref, string styno, string fabric, string content, string counts, string gg, string gsm, string Color, string bytes)
        //{
        //    string ins = "INSERT INTO ASPTBLBUYSAM (BUYERCODE ,STYLENAME ,SUBSTYLE ,FABRIC,CONTENT,COUNTS,GAUGE ,GSM ,Color,GARMENTIMAGE) VALUES('" + buyercode + "','" + styref + "','" + styno + "','" + fabric + "','" + content + "','" + counts + "','" + gg + "','" + gsm + "','" + Color + "','" + bytes + "')";
        //    Utility.ExecuteNonQuery(ins);
        //}

        public BuyerSample(string buyercode, string styref, string styno, string fabric, string content, string counts, string gg, string gsm, string Color, string bytes, Int64 ASPTBLBUYSAMID)
        {
            //string ins = "INSERT INTO TBLEMPMASIMAGE  (IMAGEBYTE,TBLEMPMASID) VALUES(:pic,'" + dt.Rows[0]["TBLEMPMASID"].ToString() + "')";
            //cmd = new MySqlCommand(ins);
            //cmd.Parameters.AddWithValueWithValue(":pic", c.ImageByte);
            //cmd.Connection = Utility.con;
            //cmd.ExecuteNonQuery();

            //string query = "UPDATE   ASPTBLBUYSAM SET BUYERCODE='" + buyercode + "',STYLENAME='" + styref + "',SUBSTYLE='" + styno + "',FABRIC='" + fabric + "', CONTENT='" + content + "',COUNTS='" + counts + "',  GAUGE='" + gg + "',GSM='" + gsm + "',Color='" + Color + "',GARMENTIMAGE='" + bytes + "' WHERE ASPTBLBUYSAMID=" + Convert.ToInt32(qrid);
            //Utility.ExecuteNonQuery(query);
        }
        public  string CapitalizeFirstLetters(string sValue)
        {



            char[] array = sValue.ToCharArray();

            if (array.Length >= 1)
            {
                if (char.IsLower(array[0]))
                {
                    array[0] = char.ToUpper(array[0]);
                }
            }

            for (int i = 1; i < array.Length; i++)
            {
                if (array[i] == ' ')
                {
                    if (char.IsLower(array[i + 1]))
                    {
                        array[i + 1] = char.ToUpper(array[i + 1]); // space is identified at Index i so next i+1 has to converted to Upper!!
                    }
                }
            }


            return new string(array);

        }
    }
}

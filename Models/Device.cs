using System.Data;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
namespace Pinnacle.Models
{
    class Device
    {
        public DataTable FromIp(string s)
        {
            string sel = ""; DataTable dt1;
            
                sel = "SELECT  distinct C.MACIP  FROM  ASPTBLMACHINEMAS A    JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE    JOIN ASPTBLMACIP C ON C.ASPTBLMACIPID=A.IPADDRESS   AND C.ACTIVE='T'    JOIN  ASPTBLUSERMAS D ON D.USERID=A.USERNAME    WHERE  B.COMPCODE='" + s + "'   ORDER BY 1";
                DataSet ds = Utility.ExecuteSelectQuery(sel, "ASPTBLMACIP");
                dt1 = ds.Tables["ASPTBLMACIP"];
            if (dt1==null)
            {
                sel = "SELECT DISTINCT  '' AS MACIP FROM DUAL UNION SELECT   B.MACIP FROM HRMACIPENTRY A JOIN HRMACIPENTRYDET B ON A.HRMACIPENTRYID = B.HRMACIPENTRYID  JOIN GTCOMPMAST C ON C.GTCOMPMASTID = A.COMPCODE  WHERE  B.CURMAC = 'YES' AND C.COMPCODE= '" + s + "'  ORDER BY  MACIP DESC";
                DataSet ds1 = Utility.ExecuteSelectQuery(sel, "HRMACIPENTRY");
                dt1 = ds1.Tables["HRMACIPENTRY"];
            }

            return dt1;

        }
        public DataTable FromIp()
        {
            string sel = ""; DataTable dt1;
            
                sel = "SELECT C.ASPTBLMACIPID, C.MACIP  FROM  ASPTBLMACHINEMAS A    JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE    JOIN ASPTBLMACIP C ON C.ASPTBLMACIPID=A.IPADDRESS   AND C.ACTIVE='T'    JOIN  ASPTBLUSERMAS D ON D.USERID=A.username    WHERE  B.COMPCODE='" + Class.Users.HCompcode + "' AND  D.USERNAME='" + Class.Users.HUserName + "'   ORDER BY 2";
                DataSet ds1 = Utility.ExecuteSelectQuery(sel, "ASPTBLMACIP");
                dt1 = ds1.Tables["ASPTBLMACIP"];
            if (dt1 == null)
            {
                sel = "SELECT B.HRMACIPENTRYDETID,B.MACIP FROM HRMACIPENTRY A JOIN HRMACIPENTRYDET B ON A.HRMACIPENTRYID = B.HRMACIPENTRYID  JOIN GTCOMPMAST C ON C.GTCOMPMASTID = A.COMPCODE  WHERE B.DEFAULTYN = 'NO' AND B.CURMAC = 'YES'  ORDER BY 2";
                DataSet ds2 = Utility.ExecuteSelectQuery(sel, "HRMACIPENTRY");
                dt1 = ds2.Tables["HRMACIPENTRY"];
            }
            return dt1;

        }
        public DataTable AllIp()
        {
            string sel = ""; DataTable dt1;

            sel = "SELECT distinct C.MACIP  FROM  ASPTBLMACHINEMAS A    JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE    JOIN ASPTBLMACIP C ON C.ASPTBLMACIPID=A.IPADDRESS   AND C.ACTIVE='T'    JOIN  ASPTBLUSERMAS D ON D.USERID=A.username      ORDER BY 1";
            DataSet ds1 = Utility.ExecuteSelectQuery(sel, "ASPTBLMACIP");
            dt1 = ds1.Tables["ASPTBLMACIP"];
            if (dt1 == null)
            {

                sel = "SELECT distinct B.MACIP FROM HRMACIPENTRY A JOIN HRMACIPENTRYDET B ON A.HRMACIPENTRYID = B.HRMACIPENTRYID  JOIN GTCOMPMAST C ON C.GTCOMPMASTID = A.COMPCODE  WHERE B.DEFAULTYN = 'NO' AND B.CURMAC = 'YES'  ORDER BY 1";
                DataSet ds2 = Utility.ExecuteSelectQuery(sel, "HRMACIPENTRY");
                dt1 = ds2.Tables["HRMACIPENTRY"];
            }
            return dt1;


        }
        public DataTable AllIp(string s, string ss)
        {
            string sel = ""; DataTable dt1;

            sel = " SELECT C.ASPTBLMACIPID, C.MACIP,c.MACIP,c.MACNO,c.MTYPE,c.MTYPE2  FROM  ASPTBLMACHINEMAS A    JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE    JOIN ASPTBLMACIP C ON C.ASPTBLMACIPID=A.IPADDRESS   AND C.ACTIVE='T'    JOIN  ASPTBLUSERMAS D ON D.USERID=A.username    and C.MACIP IN( SELECT D.MACIP FROM ASPTBLMACIP D WHERE D.MACIP='" + ss + "' ) WHERE  B.COMPCODE='" + s + "' AND  D.USERNAME='" + Class.Users.HUserName + "'   ORDER BY 2";
            DataSet ds1 = Utility.ExecuteSelectQuery(sel, "ASPTBLMACIP");
            dt1 = ds1.Tables["ASPTBLMACIP"];
            if (dt1 == null)
            {
                sel = "SELECT B.HRMACIPENTRYDETID,B.MACIP FROM HRMACIPENTRY A JOIN HRMACIPENTRYDET B ON A.HRMACIPENTRYID = B.HRMACIPENTRYID  JOIN GTCOMPMAST C ON C.GTCOMPMASTID = A.COMPCODE  WHERE B.DEFAULTYN = 'NO' AND B.CURMAC = 'YES' AND C.COMPCODE='" + s + "'  AND B.MACIP NOT IN( SELECT D.MACIP FROM HRMACIPENTRYDET D WHERE D.MACIP='" + ss + "' ) ORDER BY 2";
                DataSet ds2 = Utility.ExecuteSelectQuery(sel, "HRMACIPENTRY");
                dt1 = ds2.Tables["HRMACIPENTRY"];
            }

            return dt1;

        }
        public DataTable IPLOAD(string s, string ss)
        {
            string sel = ""; DataTable dt1;
            
                sel = "SELECT C.ASPTBLMACIPID, C.MACIP,c.MACIP,c.MACNO,c.MTYPE,c.MTYPE2  FROM  ASPTBLMACHINEMAS A    JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE     JOIN ASPTBLMACIP C ON C.ASPTBLMACIPID=A.IPADDRESS   AND C.ACTIVE='T'    JOIN  ASPTBLUSERMAS D ON D.USERID=A.username    WHERE  B.COMPCODE='" + s + "' AND  D.USERNAME='" + Class.Users.HUserName + "' and C.MACIP='" + ss + "'   ORDER BY 2";
                DataSet ds1 = Utility.ExecuteSelectQuery(sel, "ASPTBLMACIP");
                dt1 = ds1.Tables["ASPTBLMACIP"];
            if (dt1 == null)
            {
                sel = "SELECT B.MACIP,B.MACNO,B.MTYPE,B.MTYPE2 FROM HRMACIPENTRY A JOIN HRMACIPENTRYDET B ON A.HRMACIPENTRYID = B.HRMACIPENTRYID JOIN GTCOMPMAST C ON C.GTCOMPMASTID = A.COMPCODE  WHERE  B.DEFAULTYN = 'NO' AND B.CURMAC = 'YES'   AND C.COMPCODE='" + ss + "' AND B.MACIP='" + ss + "'";
                DataSet ds2 = Utility.ExecuteSelectQuery(sel, "HRMACIPENTRYDET");
                dt1 = ds2.Tables["HRMACIPENTRYDET"];

            }

            return dt1;

        }
        public DataTable ToIp(string s)
        {
            string sel = ""; DataTable dt1;
            
                sel = "SELECT C.ASPTBLMACIPID AS HRMACIPENTRYDETID, C.MACIP,c.MACIP,c.MACNO,c.MTYPE,c.MTYPE2  FROM  ASPTBLMACHINEMAS A    JOIN   GTCOMPMAST B ON B.GTCOMPMASTID = A.COMPCODE    JOIN ASPTBLMACIP C ON C.ASPTBLMACIPID=A.IPADDRESS   AND C.ACTIVE='T'    JOIN  ASPTBLUSERMAS D ON D.USERID=A.username  WHERE  B.COMPCODE='" + s + "' AND  D.USERNAME='" + Class.Users.HUserName + "'   ORDER BY 2";
                DataSet ds1 = Utility.ExecuteSelectQuery(sel, "ASPTBLMACIP");
                dt1 = ds1.Tables["ASPTBLMACIP"];
            if (dt1 == null)
            {
                sel = "SELECT B.HRMACIPENTRYDETID,B.MACIP FROM HRMACIPENTRY A JOIN HRMACIPENTRYDET B ON A.HRMACIPENTRYID = B.HRMACIPENTRYID  JOIN GTCOMPMAST C ON C.GTCOMPMASTID = A.COMPCODE  WHERE B.DEFAULTYN = 'NO' AND B.CURMAC = 'YES' AND C.COMPCODE='" + s + "'  ORDER BY 2";
                DataSet ds2 = Utility.ExecuteSelectQuery(sel, "HRMACIPENTRYDET");
                dt1 = ds2.Tables["HRMACIPENTRYDET"];
            }

            return dt1;
        }
        public DataTable fingerindex(string s)
        {
            string sel = "SELECT distinct  a.FINGER_INDEX FROM  TFTDevice a where A.CURMAC='T'  AND a.user_id='" + s + "' order by 1 ";
            DataSet ds1 = Utility.ExecuteSelectQuery(sel, "TFTDevice");
            DataTable dt1 = ds1.Tables["TFTDevice"];
            return dt1;
        }

        public DataTable userid()
        {
            string sel = "SELECT distinct a.USER_ID FROM  TFTDevice a  order by 1 ";
            DataSet ds1 = Utility.ExecuteSelectQuery(sel, "TFTDevice");
            DataTable dt1 = ds1.Tables["TFTDevice"];
            return dt1;
        }
        public static byte[] ImageToByteArray(PictureBox imageIn)
        {
            var ms = new MemoryStream();
            imageIn.Image.Save(ms, imageIn.Image.RawFormat);
            return ms.ToArray();
        }
        public Image ByteArrayToImage1(byte[] byteArrayIn)
        {
            System.Drawing.ImageConverter converter = new System.Drawing.ImageConverter();
            Image img = (Image)converter.ConvertFrom(byteArrayIn);

            return img;
        }
        public static Image ByteArrayToImage(byte[] byteArrayIn)
        {
            var ms = new MemoryStream(byteArrayIn);
            Image returnImage = Image.FromStream(ms);
            return returnImage;
        }

        public static string BytesToString(byte[] bytes)
        {
            using (MemoryStream stream = new MemoryStream(bytes))
            {
                using (StreamReader streamReader = new StreamReader(stream))
                {
                    return streamReader.ReadToEnd();
                }
            }
        }

        //internal static byte[] ImageToByteArray(Bitmap img)
        //{

        //}

        //public byte[] ImageAArray(System.Drawing.Image imagen)
        //{
        //    System.IO.MemoryStream ms = new System.IO.MemoryStream();
        //    imagen.Save(ms, System.Drawing.Imaging.ImageFormat.Gif);
        //    return ms.ToArray();
        //}
        //public System.Drawing.Image ArrayAImage(byte[] ArrBite)
        //{
        //    System.IO.MemoryStream ms = new System.IO.MemoryStream(ArrBite);
        //    System.Drawing.Image returnImage = System.Drawing.Image.FromStream(ms);
        //    return returnImage;
        //}
    }
   
}

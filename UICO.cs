using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MySql.Data.MySqlClient;
using System.Data;

namespace Pinnacle
{
    class UICO
    {
       // public MySqlConnection asconn;
        MySqlCommand ascmd;
        DataTable dt;
        MySqlDataAdapter da;
       
     
        public void Update_User_DetailsTFT(string UserId, string Name, int fingerIndex,string card, int privilege, string pass, bool chk, int flag, string macip)
        {
       
            //try
            //{
                
              
            //    string ch = "";
            //    if (chk == true) { ch = "T"; } else { ch = "F"; }
            //    string query = "UPDATE   TFTDevice SET User_Id=:User_Id,Name=':Name',  Finger_Index=:Finger_Index,cardnumber=:cardnumber,Privilege=:Privilege, Passwords=:Passwords,ENABLED=:ENABLED,Flag=:flag ,MACIP=:MacIP WHERE User_Id=:User_Id AND Name=:Name  AND MACIP=:MacIP";
            //    ascmd = new MySqlCommand(query,Utility.con);
            //    ascmd.Parameters.AddWithValue(":User_Id", UserId);
            //    ascmd.Parameters.AddWithValue(":Name", Name);
            //    ascmd.Parameters.AddWithValue(":Finger_Index", fingerIndex);              
            //    ascmd.Parameters.AddWithValue(":cardnumber", card);
            //    ascmd.Parameters.AddWithValue(":Privilege", privilege);
            //    ascmd.Parameters.AddWithValue(":Passwords", pass);
            //    ascmd.Parameters.AddWithValue(":ENABLED", ch);
            //    ascmd.Parameters.AddWithValue(":flag", flag);
            //    ascmd.Parameters.AddWithValue(":MacIP", macip);

            //    ascmd.ExecuteNonQuery();

            //}
            //catch (Exception ex)
            //{
            //    ////throw ex;
            //}
            //finally
            //{
            //    Utility.con.Close();
            //}

        }
       
        public void Update_User_DetailsTFT1(string UserId, string Name, int fingerIndex, string fingerImage, string Face_Image, string card, int privilege, string pass, bool chk, int flag, string macip)
        {
            //try
            //{
               
            //    string ch = "";
            //    if (chk == true) { ch = "T"; } else { ch = "F"; }
            //    byte[] bytes = Encoding.UTF8.GetBytes(Convert.ToString(Face_Image));
            //    byte[] finbytes = Encoding.UTF8.GetBytes(Convert.ToString(fingerImage));
            //    string query = "UPDATE   TFTDevice SET User_Id=:User_Id,Name=:Name,  Finger_Index=:Finger_Index,Finger_Image=:Finger_Image,Face_Image=:Face_Image,cardnumber=:cardnumber,Privilege=:Privilege, Passwords=:Passwords,ENABLED=:ENABLED,Flag=:flag ,MACIP=:MacIP,FACEIMAGE=:faceimage,FINGER_IMAGE=:fingerlength,FACELENGTH=:facelength,FINGERIMAGE=:FINGERIMAGE WHERE User_Id=:User_Id  ";
            //    ascmd = new MySqlCommand(query,Utility.con);
            //    ascmd.Parameters.AddWithValue(":User_Id", UserId);
            //    ascmd.Parameters.AddWithValue(":Name", Name);
            //    ascmd.Parameters.AddWithValue(":Finger_Index", fingerIndex);
            //    ascmd.Parameters.AddWithValue(":Finger_Image", fingerImage);
            //    ascmd.Parameters.AddWithValue(":Face_Image", Face_Image);
            //    ascmd.Parameters.AddWithValue(":cardnumber", card);
            //    ascmd.Parameters.AddWithValue(":Privilege", privilege);
            //    ascmd.Parameters.AddWithValue(":Passwords", pass);
            //    ascmd.Parameters.AddWithValue(":ENABLED", ch);
            //    ascmd.Parameters.AddWithValue(":flag", flag);
            //    ascmd.Parameters.AddWithValue(":MacIP", macip);
            //    ascmd.Parameters.AddWithValue(":faceimage", bytes);
            //    ascmd.Parameters.AddWithValue(":fingerlength", fingerImage.Length);
            //    ascmd.Parameters.AddWithValue(":facelength", Face_Image.Length);
            //    ascmd.Parameters.AddWithValue(":FINGERIMAGE", finbytes);
            //    ascmd.ExecuteNonQuery();
            //}
            //catch (Exception ex)
            //{
            //    ////throw ex;
            //}
            //finally
            //{
            //    Utility.con.Close();
            //}

        }
        public void Insert_User_DetailsTFT2(string UserId, string Name, int fingerIndex, string fingerImage, string Face_Image, string card, int privilege, string pass, bool chk, int flag,string macip)
        {
            //try
            //{
               
            //    //string ins = "INSERT INTO TFTDevice(USER_ID,NAME,FINGER_INDEX,CARDNUMBER,PRIVILEGE,PASSWORDS,ENABLED,FLAG,MACIP) VALUES('"+UserId+ "','" + Name + "'," + fingerIndex + ",'" + card + "','" + privilege + "','" + pass + "','" + chk + "'," + flag + ",'" + macip + "')";
            //    // Utility.ExecuteNonQuery(ins);
            //    byte[] bytes = Encoding.UTF8.GetBytes(Convert.ToString(Face_Image));
            //    byte[] finbytes = Encoding.UTF8.GetBytes(Convert.ToString(fingerImage));
            //    string ch = "";
            //    if (chk == true) { ch = "T"; } else { ch = "F"; }
            //    string ins = "INSERT INTO TFTDevice(USER_ID,NAME,FINGER_INDEX,FINGER_IMAGE,FACE_IMAGE,CARDNUMBER,PRIVILEGE,PASSWORDS,ENABLED,FLAG,MACIP,FACEIMAGE,FINGER_LENGTH,FACELENGTH,FINGERIMAGE) VALUES(:user_Id,:name,:finger_Index,:finger_Image,:face_Image,:card,:privilege,:passwords,:enabled,:flag,:macIP,:faceimage,:fingerlength,:facelength,:fingerimage)";
            //    ascmd = new MySqlCommand(ins, Utility.con);
            //    ascmd.Parameters.AddWithValue(":user_Id", UserId);
            //    ascmd.Parameters.AddWithValue(":name", Name);
            //    ascmd.Parameters.AddWithValue(":finger_Index", fingerIndex);
            //    ascmd.Parameters.AddWithValue(":finger_Image", fingerImage);
            //    ascmd.Parameters.AddWithValue(":face_Image", Face_Image);
            //    ascmd.Parameters.AddWithValue(":card", card);
            //    ascmd.Parameters.AddWithValue(":privilege", privilege);
            //    ascmd.Parameters.AddWithValue(":passwords", pass);
            //    ascmd.Parameters.AddWithValue(":enabled", ch);
            //    ascmd.Parameters.AddWithValue(":flag", flag);
            //    ascmd.Parameters.AddWithValue(":macIP", macip);
            //    ascmd.Parameters.AddWithValue(":faceimage", bytes);
            //    ascmd.Parameters.AddWithValue(":fingerlength", fingerImage.Length);
            //    ascmd.Parameters.AddWithValue(":facelength", Face_Image.Length);
            //    ascmd.Parameters.AddWithValue(":fingerimage", finbytes);

            //    ascmd.ExecuteNonQuery();

                

            //}
            //catch (Exception ex)
            //{
            //    ////throw ex;
            //}
            //finally
            //{
            //     Utility.con.Close();
            //}

        }

     
        public DataTable UploadDataTFT()
        {

            //try
            //{
             
            //    string sel = "SELECT * FROM  TFTDevice";
            //    DataSet ds1 = Utility.ExecuteSelectQuery(sel, "TFTDevice");
            //    DataTable dt = ds1.Tables["TFTDevice"];
            //    return dt;
               
            //}
            //catch (Exception ex)
            //{
            //    ////throw ex;
            //}
            //finally
            //{
            //   // Utility.con.Close();
            //}
            return dt;

        }
        public DataTable UploadDataTFT(string ip)
        {

            try
            {



                string sel = "SELECT * FROM  " + Class.Users.HCompcode + "TRS_ATTLOG where MACIP='" + ip + "'";
                DataSet ds1 = Utility.ExecuteSelectQuery(sel, "TRS_ATTLOG");
                DataTable dt = ds1.Tables["TRS_ATTLOG"];
                return dt;


            }
            catch (Exception ex)
            {
                //throw ex;
            }
           
            return dt;

        }
        public DataTable UploadDataTFT(string userid,int Finger_Index)//,int Finger_Index, string sCardnumber
        {

            try
            {
                //// byte[] bytes1 = Encoding.UTF8.GetBytes(Convert.ToString(fingerimage));
                //   byte[] bytes2 = Encoding.UTF8.GetBytes(Convert.ToString(faceimage)); AND FINGER_LENGTH='" + bytes1.Length + "' AND FACELENGTH='" + bytes2.Length + "'
                string sel = "SELECT  TFTDeviceID FROM  TFTDevice where  User_Id='" + userid + "' AND Finger_Index=" + Finger_Index;// AND   Finger_Index=" + Finger_Index + " AND cardnumber='" + sCardnumber + "' 
                DataSet ds1 = Utility.ExecuteSelectQuery(sel, "TFTDevice");
                DataTable dt = ds1.Tables["TFTDevice"];
                return dt;


            }
            catch (Exception ex)
            {
                //throw ex;
            }
            finally
            {
               // Utility.con.Close();
            }
            return dt;

        }

        public DataTable UploadDataTFT(string userid, string name, string ip,string chk)
        {

            try
            {

                string sel = "SELECT User_Id,Finger_Index,Enabled FROM  TFTDevice where  User_Id='" + userid + "' AND   NAME='" + name + "'   AND MACIP='" + ip + "' AND ENABLED='" + chk + "' ";
                DataSet ds1 = Utility.ExecuteSelectQuery(sel, "TFTDevice");
                DataTable dt = ds1.Tables["TFTDevice"];
                return dt;


            }
            catch (Exception ex)
            {
                //throw ex;
            }
            finally
            {
                // Utility.con.Close();
            }
            return dt;

        }
        public void DeleteAllEmpTmTFT()
        {
            try
            {
              
                string del = "Delete from TFTDevice ";
                Utility.ExecuteNonQuery(del);


            }
            catch (Exception ex)
            {
                 //throw ex;
            }
            finally
            {
                Utility.con.Close();
            }
        }

        public void DeleteAllEmpTmTFT(string userid,int backno)
        {
            try
            {
               
                string s = "Delete from TFTDevice where User_Id='" + userid + "' and finger_index='" + backno + "' ";
                Utility.ExecuteNonQuery(s);
               

            }
            catch (Exception ex)
            {
                //throw ex;
            }
            finally
            {
                Utility.con.Close();
            }
        }

        public void DeleteAllEmpTmTFT(int backno)
        {
            try
            {
               
                string s = "Delete from TFTDevice where  finger_index='" + backno + "' ";
                Utility.ExecuteNonQuery(s);


            }
            catch (Exception ex)
            {
                //throw ex;
            }
            finally
            {
                Utility.con.Close();
            }

        }
        // for IFACE Finger templates
        public void Insert_User_DetailsIFACE_FingerTm(string UserId, string Name, int fingerIndex, string fingerImage, int privilege, string pass, bool enabled, int flag)
        {

            try
            {
               
                string chk;
                if (enabled == false) chk = "false"; else chk = "true";
                string query = "INSERT INTO IFACEDevice_FingerTm(User_Id,Name,Finger_Index,Finger_Image,Privilege,Passwords,Enabled,Flag) VALUES(':userId',':name',:fingerIndex,':fingerImage',:privilege,':password',':enabled',':flag')";

                ascmd = new MySqlCommand(query);

                ascmd.Parameters.AddWithValue(":userId", UserId);
                ascmd.Parameters.AddWithValue(":name", Name);
                ascmd.Parameters.AddWithValue(":fingerIndex", fingerIndex);
                ascmd.Parameters.AddWithValue(":fingerImage", fingerImage);
                ascmd.Parameters.AddWithValue(":privilege", privilege);
                ascmd.Parameters.AddWithValue(":password", pass);
                ascmd.Parameters.AddWithValue(":enabled", chk);
                ascmd.Parameters.AddWithValue(":flag", flag);

                ascmd.Connection = Utility.con;
                ascmd.ExecuteNonQuery();

               

            }
            catch (Exception ex)
            {
                //throw ex;
            }
            finally
            {
                Utility.con.Close();
            }

        }
        public DataTable UploadDataIFACE_FingerTm(string macip)
        {

            try
            {
                string sel = "SELECT * FROM  IFACEDevice_FaceTm where MACIP='" + macip + "' ";
                DataSet ds1 = Utility.ExecuteSelectQuery(sel, "IFACEDevice_FaceTm");
                DataTable dt = ds1.Tables["IFACEDevice_FaceTm"];
                return dt;
            }
            catch (Exception ex)
            {
                //throw ex;
            }
            finally
            {
                Utility.con.Close();
            }
            return dt;

        }
        public void DeleteAllEmpTmIFACE_FingerTm()
        {
            try
            {
              
                string del = "Delete from IFACEDevice_FingerTm ";
                Utility.ExecuteNonQuery(del);

            }
            catch (Exception ee)
            {
                // throw ee;
            }
            finally
            {
                Utility.con.Close();
            }
        }

        // for IFACE Face templates
        public void Insert_User_DetailsIFACE_FaceTm(string UserId, string Name, string pass, int privilege, int faceIndex, string faceImage, int faceLength, bool enabled,string macip)
        {


            try
            {

                byte[] bytes = Encoding.UTF8.GetBytes(Convert.ToString(faceImage));

                //string query = "INSERT INTO IFACEDevice_FaceTm(User_Id,Name,Passwords,Privilege,Face_Index,Face_Length,ENABLED,MacIP) VALUES('" + UserId + "','" + Name + "','" + pass + "'," + privilege + "," + faceIndex + "," + faceLength + ",'" + enabled + "','" + macip + "')";
                //Utility.ExecuteNonQuery(query);

                string ins = "INSERT INTO IFACEDevice_FaceTm (User_Id,Name,Passwords,Privilege,Face_Index,Face_Image,Face_Length,ENABLED,MacIP,FACEIMAGE) VALUES(:User_Id,:Name,:Passwords,:Privilege,:Face_Index,:Face_Image,:Face_Length,:ENABLED,:MacIP,:FACEIMAGE)";
                ascmd = new MySqlCommand(ins, Utility.con);
                ascmd.Parameters.AddWithValue(":User_Id", UserId);
                ascmd.Parameters.AddWithValue(":Name", Name);
                ascmd.Parameters.AddWithValue(":Passwords", pass);
                ascmd.Parameters.AddWithValue(":Privilege", privilege);
                ascmd.Parameters.AddWithValue(":Face_Index", faceIndex);
                ascmd.Parameters.AddWithValue(":Face_Image", faceImage);
                ascmd.Parameters.AddWithValue(":Face_Length", faceLength);
                ascmd.Parameters.AddWithValue(":ENABLED", enabled);
                ascmd.Parameters.AddWithValue(":MacIP", macip);
                ascmd.Parameters.AddWithValue(":FACEIMAGE", bytes);
                ascmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                //throw ex;
            }
            finally
            {
                Utility.con.Close();
            }

        }
        public DataTable UploadDataIFACE_FaceTm()
        {

            try
            {
                dt = new DataTable();
               

                string upload_data = "SELECT * FROM  IFACEDevice_FaceTm";
                ascmd = new MySqlCommand(upload_data, Utility.con);
                da = new MySqlDataAdapter(ascmd);
                da.Fill(dt);

            }
            catch (Exception ex)
            {
                //throw ex;
            }
            finally
            {
                Utility.con.Close();
            }
            return dt;

        }
        public void DeleteAllEmpTmIFACE_FaceTm()
        {
            try
            {
              
                string CheckValPersonId = "Delete from IFACEDevice_FaceTm ";
                ascmd = new MySqlCommand(CheckValPersonId, Utility.con);
                ascmd.ExecuteNonQuery();


            }
            catch (Exception ee)
            {
                 throw ee;
            }
            finally
            {
                Utility.con.Close();
            }
        }

        internal void insertimage(string sTmpData, string sName)
        {

            byte[] ImageByte = Encoding.ASCII.GetBytes(sTmpData);


            //    string query = "INSERT INTO [TBLEMPMASIMAGE](IMAGEBYTE) VALUES(@faceImage)";

            MySqlCommand cmd;
            string ins = "INSERT INTO TBLEMPMASIMAGE  (IMAGEBYTE,image,username) VALUES(:pic,:pic1,:username)";
            cmd = new MySqlCommand(ins);
            cmd.Parameters.AddWithValue(":pic", ImageByte);
            cmd.Parameters.AddWithValue(":pic1", sTmpData);
            cmd.Parameters.AddWithValue(":username", sName);
            cmd.Connection = Utility.con;
            cmd.ExecuteNonQuery();

        }
    }
}

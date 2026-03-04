using System;
using System.Collections.Generic;
using System.Text;
using System.Configuration;
using System.Collections;
using System.Windows.Forms;
using System.Data;
using MySql.Data.MySqlClient;
using Pinnacle.UserControls;
using System.Data.OracleClient;
namespace Pinnacle
{

   
    public class Utility
    {
      
        public static MySqlConnection con;
        private static MySqlCommand cmd;
        public static MySqlTransaction Trans;
        public static MySqlDataReader Rdr;
        public static MySqlDataAdapter da;

        static Utility()
        {
            try
            {


                try
                {
                    string[] data = Class.Users.ConString.Split(',');
                    if (con != null)
                    {
                        if (con.State == ConnectionState.Closed)
                        {
                            con = new MySqlConnection(data[0]);
                            con.Open();
                        }
                    }
                    else
                    {
                        con = new MySqlConnection(data[0]);
                        con.Open();
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }
    
        public static MySqlConnection Connect()
        {
            try
            {
                if (con.State == ConnectionState.Closed)
                {
                    con.Open();
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            return con;
        }
        public static MySqlConnection DisConnect()
        {
            //try
            //{
            //    if (con.State == ConnectionState.Open)
            //    {
            //        con.Close();

            //    }

            //}
            //catch (Exception ex)
            //{
            //    throw ex;
            //}
            return con;
        }
        public DataTable select(string qry, string tbl)
        {
            DataSet ds = Utility.ExecuteSelectQuery(qry, tbl);
            DataTable dt = ds.Tables[tbl];
            return dt;
        }
        public static DataSet ExecuteSelectQuery(
      string query,
      string tableName,
      Dictionary<string, object> parameters = null)
        {
            var ds = new DataSet();

            try
            {
                using (var con = Connect()) // always new connection
                using (var cmd = new MySqlCommand(query, con))
                {
                    cmd.CommandType = CommandType.Text;

                    if (parameters != null)
                    {
                        foreach (var param in parameters)
                        {
                            cmd.Parameters.AddWithValue(
                                param.Key.StartsWith("@") ? param.Key : "@" + param.Key,
                                param.Value ?? DBNull.Value
                            );
                        }
                    }

                    using (var da = new MySqlDataAdapter(cmd))
                    {
                        da.Fill(ds, tableName);
                    }
                }
            }
            catch (MySqlException ex)
            {
                // Log and rethrow (preferred)
                throw new ApplicationException("Database operation failed", ex);
            }

            return ds;
        }


        public static DataSet ExecuteSelectQuery(string query, string tblname)
        {
            DataSet ds = new DataSet();
            try
            {
                if (con.State == ConnectionState.Closed)
                {
                    Connect();
                }
                da = new MySqlDataAdapter(query, con);
                da.Fill(ds, tblname);

            }
            catch { }
            DisConnect();
            return ds;
        }
        public static void  ExecuteNonQuery(string query)
        {
            cmd = new MySqlCommand(query, con);
            cmd.Transaction = Trans;
            cmd.CommandTimeout = 360;
            cmd.ExecuteNonQuery();
        }
        public static bool ExecuteNonQuery(string query, Byte[] s)
        {
            if (con.State == ConnectionState.Closed)
            {
                Connect();
            }
            cmd.CommandTimeout = 360;
            cmd = new MySqlCommand(query, con);           
            cmd.Parameters.AddWithValue(":EMPIMAGE", s);
            cmd.ExecuteNonQuery();
           
            DisConnect();
            return true;
        }
        public static MySqlDataReader ExecuteReader(string query)
        {
            if (con.State == ConnectionState.Closed)
            {
                Connect();
            }
            cmd.CommandTimeout = 360;
            cmd = new MySqlCommand(query, con);
            MySqlDataReader dr = cmd.ExecuteReader();
            DisConnect();
            return dr;
        }
        public static object ExecuteScalar(string sql)
        {
            if (con.State == ConnectionState.Closed)
            {
                Connect();
            }
            cmd.CommandTimeout = 360;
            cmd = new MySqlCommand(sql, con);
            object scalarValue = cmd.ExecuteScalar();
            DisConnect();
            return scalarValue;

        }
        public static void Load_ListCombo(object Sender, string Sql, string ValMem, string DisMem, Hashtable ParamTable = null, string DefValue = "")
        {
            // ''''' Load Combo / Listbox
            // Warning!!! Optional parameters not supported
            try
            {
                MySqlCommand CmdData = new MySqlCommand();
                con.Open();
                CmdData.CommandText = Sql;
                CmdData.Connection = con;
                if (ParamTable != null)
                {
                    foreach (DictionaryEntry DeData in ParamTable)
                    {
                        CmdData.Parameters.AddWithValue(DeData.Key.ToString(), DeData.Value);
                    }

                }

                MySqlDataAdapter Sda = new MySqlDataAdapter(CmdData);
                DataTable Dt = new DataTable();
                Sda.Fill(Dt);
                if (DefValue != "")
                {
                    DataRow Row;
                    Row = Dt.NewRow();
                    Row[ValMem] = 0;
                    Row[DisMem] = DefValue;
                    Dt.Rows.InsertAt(Row, 0);
                }
                ((UCComboBox)Sender).SQLQuery = Sql;
                ((UCComboBox)Sender).DataSource = null;
                ((UCComboBox)Sender).DisplayMember = DisMem;
                ((UCComboBox)Sender).ValueMember = ValMem;
                ((UCComboBox)Sender).DataSource = Dt;
                if (((UCComboBox)Sender).DropDownStyle == ComboBoxStyle.DropDown)
                {
                    if (DefValue != "")
                    {
                        ((UCComboBox)Sender).SelectedIndex = 0;
                    }
                    else
                    {
                        ((UCComboBox)Sender).SelectedIndex = -1;
                    }
                }
            }
            catch (Exception ex)
            {
            }
            finally
            {
                DisConnect();
            }

        }
        public static void Load_GridCombo(object Sender, string Sql, string ValMem, string DisMem, Hashtable ParamTable = null, string DefValue = "")
        {
            // ''''' Load Combo / Listbox
            // Warning!!! Optional parameters not supported
            try
            {
                MySqlCommand CmdData = new MySqlCommand();
                con.Open();
                CmdData.CommandText = Sql;
                CmdData.Connection = con;
                if (ParamTable != null)
                {
                    foreach (DictionaryEntry DeData in ParamTable)
                    {
                        CmdData.Parameters.AddWithValue(DeData.Key.ToString(), DeData.Value);
                    }

                }

                MySqlDataAdapter Sda = new MySqlDataAdapter(CmdData);
                DataTable Dt = new DataTable();
                Sda.Fill(Dt);
                if (DefValue != "")
                {
                    DataRow Row;
                    Row = Dt.NewRow();
                    Row[ValMem] = 0;
                    Row[DisMem] = DefValue;
                    Dt.Rows.InsertAt(Row, 0);
                }               
                ((DataGridViewComboBoxColumn)Sender).DataSource = null;
                ((DataGridViewComboBoxColumn)Sender).DisplayMember = DisMem;
                ((DataGridViewComboBoxColumn)Sender).ValueMember = ValMem;
                ((DataGridViewComboBoxColumn)Sender).DataSource = Dt;
                if (((ComboBox)Sender).DropDownStyle == ComboBoxStyle.DropDown)
                {
                    if (DefValue != "")
                    {
                        ((ComboBox)Sender).SelectedIndex = 0;
                    }
                    else
                    {
                        ((ComboBox)Sender).SelectedIndex = -1;
                    }
                }
            }
            catch (Exception ex)
            {
            }
            finally
            {
                DisConnect();
            }

        }
        public static void Load_Combo(object Sender, string Sql, string ValMem, string DisMem, Hashtable ParamTable = null, string DefValue = "", string DefValueOrder = "TOP", bool QueryUpdate = true)
        {
            // ''''' Load Combo / Listbox
            // Warning!!! Optional parameters not supported
            try
            {
                
                MySqlCommand CmdData = new MySqlCommand();
                if (con.State == ConnectionState.Closed)
                {
                    Connect();
                }
                CmdData.CommandText = Sql;
                CmdData.Connection = con;
                if (ParamTable != null)
                {
                    foreach (DictionaryEntry DeData in ParamTable)
                    {
                        CmdData.Parameters.AddWithValue(DeData.Key.ToString(), DeData.Value);
                    }

                }

                MySqlDataAdapter Sda = new MySqlDataAdapter(CmdData);
                DataTable Dt = new DataTable();
                Sda.Fill(Dt);
                if (DefValue != "")
                {
                    DataRow Row;
                    Row = Dt.NewRow();
                    Row[ValMem] = 0;
                    Row[DisMem] = DefValue;
                    if (DefValueOrder.ToUpper() == "BOTTOM")
                    {
                        if (Dt.Rows.Count + 1 > 0)
                        {
                            Dt.Rows.InsertAt(Row, Dt.Rows.Count + 1);
                        }
                        else
                        {
                            Dt.Rows.InsertAt(Row, 0);
                        }
                    }
                    else
                    {
                        Dt.Rows.InsertAt(Row, 0);
                    }
                }               
                ((ComboBox)Sender).DisplayMember = DisMem;
                ((ComboBox)Sender).ValueMember = ValMem;
                ((ComboBox)Sender).DataSource = Dt;
                //if (((ComboBox)Sender).DropDownStyle == ComboBoxStyle.DropDown)
                //{
                //    if (DefValue != "")
                //    {
                //        ((ComboBox)Sender).SelectedIndex = 0;
                //    }
                //    else
                //    {
                //        ((ComboBox)Sender).SelectedIndex = -1;
                //    }
                //}
            }
            catch (Exception ex)
            {
            }
            finally
            {
                DisConnect();
            }

        }
        public static void Load_DataGrid(object Sender, string Sql, Hashtable ParamTable = null)
        {
            MySqlCommand CmdData = new MySqlCommand();
            // Warning!!! Optional parameters not supported
            MySqlDataAdapter Sda = new MySqlDataAdapter();
            DataSet Ds = new DataSet();
            DataView Dv = new DataView();
            try
            {
                Connect();
                CmdData.CommandText = Sql;
                CmdData.Connection = con;
                CmdData.CommandTimeout = 180;
                if (ParamTable != null)
                {
                    foreach (DictionaryEntry DeData in ParamTable)
                    {
                        CmdData.Parameters.AddWithValue(DeData.Key.ToString(), DeData.Value);
                    }

                }

                Sda = new MySqlDataAdapter(CmdData);
                Ds = new DataSet();
                Sda.Fill(Ds, "tabresult");
                ((DataGridView)(Sender)).DataSource = Ds.Tables[0];
               
            }
            catch (Exception ex)
            {
            }
            finally
            {
                DisConnect();
            }

        }
      public static DataTable SQLQuery(string Sql, Hashtable ParamTable = null)
        {
            MySqlCommand CmdData = new MySqlCommand();
            // Warning!!! Optional parameters not supported
            MySqlDataAdapter Sda = new MySqlDataAdapter();
            DataSet Ds = new DataSet();            
            Connect(); 
            CmdData.CommandText = Sql;
            CmdData.Connection = con;
            CmdData.Transaction = Trans;
            CmdData.CommandTimeout = 180;
            if (ParamTable != null)
            {
                foreach (DictionaryEntry DeData in ParamTable)
                {
                    CmdData.Parameters.AddWithValue(DeData.Key.ToString(), DeData.Value);
                }
            }

            Sda = new MySqlDataAdapter(CmdData);          
            Sda.Fill(Ds, "tabresult");
            DataTable SQLQuery = new DataTable();
            SQLQuery.Clear();
            SQLQuery = Ds.Tables["tabresult"];
            return SQLQuery;
        }

        public static void SQLNonQuery(string Sql, Hashtable ParamTable = null)
        {
            MySqlCommand CmdData = new MySqlCommand();
            // Warning!!! Optional parameters not supported
            Connect(); 
            CmdData.CommandText = Sql;
            CmdData.Connection = con;
            CmdData.Transaction = Trans;
            CmdData.CommandTimeout = 180;
            if (ParamTable != null)
            {
                foreach (DictionaryEntry DeData in ParamTable)
                {
                    CmdData.Parameters.AddWithValue(DeData.Key.ToString(), DeData.Value);
                }
            }
            CmdData.ExecuteNonQuery();
        }

        public static void SQLNonQuery1(string Sql, DataTable Dt, Hashtable ParamTable = null)
        {
            try
            {
                MySqlCommand CmdData = new MySqlCommand();
                // Warning!!! Optional parameters not supported
                Connect(); 
                CmdData.CommandText = Sql;
                CmdData.Connection = con;
                CmdData.CommandType = CommandType.StoredProcedure;
                CmdData.Transaction = Trans;
                CmdData.CommandTimeout = 180;
                if (!(ParamTable == null))
                {
                    foreach (DictionaryEntry DeData in ParamTable)
                    {
                        CmdData.Parameters.AddWithValue(DeData.Key.ToString(), DeData.Value);
                    }

                }
                CmdData.Parameters.AddWithValue("@Dt", Dt);
                CmdData.ExecuteNonQuery();
            }
            catch (Exception ex) { MessageBox.Show(ex.Message.ToString()); }

        }

        public static object SQLScalar(string Sql, Hashtable ParamTable = null)
        {
            MySqlCommand CmdData = new MySqlCommand();
            // Warning!!! Optional parameters not supported
            Connect();
            CmdData.CommandText = Sql;
            CmdData.Connection = con;
            CmdData.Transaction = Trans;
            if (ParamTable != null)
            {
                foreach (DictionaryEntry DeData in ParamTable)
                {
                    CmdData.Parameters.AddWithValue(DeData.Key.ToString(), DeData.Value);
                }

            }

            return CmdData.ExecuteScalar();
        }

        public static MySqlDataReader SQLReader(string Sql, Hashtable ParamTable = null)
        {
            MySqlCommand CmdData = new MySqlCommand();
            // Warning!!! Optional parameters not supported
            Connect(); 
            if (Rdr != null)
            {
                Rdr.Close();
            }

            CmdData.CommandText = Sql;
            CmdData.Connection = con;
            if (ParamTable != null)
            {
                foreach (DictionaryEntry DeData in ParamTable)
                {
                    CmdData.Parameters.AddWithValue(DeData.Key.ToString(), DeData.Value);
                }

            }
            Rdr = CmdData.ExecuteReader();
            return Rdr;
        }

        //public static void Load_ListCombo(object Sender, string Sql, string ValMem, string DisMem, Hashtable ParamTable = null, string DefValue = "")
        //{
        //    // ''''' Load Combo / Listbox
        //    // Warning!!! Optional parameters not supported
        //    try
        //    {
        //        MySqlCommand CmdData = new MySqlCommand();
        //        Utility.Connect(); 
        //        CmdData.CommandText = Sql;
        //        CmdData.Connection = con;
        //        if (ParamTable != null)
        //        {
        //            foreach (DictionaryEntry DeData in ParamTable)
        //            {
        //                CmdData.Parameters.AddWithValue(DeData.Key.ToString(), DeData.Value);
        //            }

        //        }

        //        MySqlDataAdapter Sda = new MySqlDataAdapter(CmdData);
        //        DataTable Dt = new DataTable();
        //        Sda.Fill(Dt);
        //        if (DefValue != "")
        //        {
        //            DataRow Row;
        //            Row = Dt.NewRow();
        //            Row[ValMem] = 0;
        //            Row[DisMem] = DefValue;
        //            Dt.Rows.InsertAt(Row, 0);
        //        }
        //        ((UCComboBox)Sender).SQLQuery = Sql;
        //        ((UCComboBox)Sender).DataSource = null;
        //        ((UCComboBox)Sender).DisplayMember = DisMem;
        //        ((UCComboBox)Sender).ValueMember = ValMem;
        //        ((UCComboBox)Sender).DataSource = Dt;
        //        if (((UCComboBox)Sender).DropDownStyle == ComboBoxStyle.DropDown)
        //        {
        //            if (DefValue != "")
        //            {
        //                ((UCComboBox)Sender).SelectedIndex = 0;
        //            }
        //            else
        //            {
        //                ((UCComboBox)Sender).SelectedIndex = -1;
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //    }
        //    finally
        //    {
        //        Utility.DisConnect(); 
        //    }

        //}

        //public static void Load_DataGrid(object Sender, string Sql, Hashtable ParamTable = null)
        //{
        //    MySqlCommand CmdData = new MySqlCommand();
        //    // Warning!!! Optional parameters not supported
        //    MySqlDataAdapter Sda = new MySqlDataAdapter();
        //    DataSet Ds = new DataSet();
        //    DataView Dv = new DataView();
        //    try
        //    {
        //        Utility.Connect();
        //        CmdData.CommandText = Sql;
        //        CmdData.Connection = con;
        //        CmdData.CommandTimeout = 180;
        //        if (ParamTable != null)
        //        {
        //            foreach (DictionaryEntry DeData in ParamTable)
        //            {
        //                CmdData.Parameters.AddWithValue(DeData.Key.ToString(), DeData.Value);
        //            }

        //        }

        //        Sda = new MySqlDataAdapter(CmdData);
        //        Ds = new DataSet();
        //        Sda.Fill(Ds, "tabresult");
        //        ((DataGridView)(Sender)).DataSource = Ds.Tables[0];
        //    }
        //    catch (Exception ex)
        //    {
        //    }
        //    finally
        //    {
        //        Utility.DisConnect(); 
        //    }

        //}

        //public static DataTable SQLQuery(string Sql, Hashtable ParamTable = null)
        //{
        //    MySqlCommand CmdData = new MySqlCommand();
        //    // Warning!!! Optional parameters not supported
        //    MySqlDataAdapter Sda = new MySqlDataAdapter();
        //    DataSet Ds = new DataSet();
        //    Utility.Connect(); ();
        //    CmdData.CommandText = Sql;
        //    CmdData.Connection = con;
        //    CmdData.Transaction = Trans;
        //    CmdData.CommandTimeout = 180;
        //    if (ParamTable != null)
        //    {
        //        foreach (DictionaryEntry DeData in ParamTable)
        //        {
        //            CmdData.Parameters.AddWithValue(DeData.Key.ToString(), DeData.Value);
        //        }

        //    }

        //    Sda = new MySqlDataAdapter(CmdData);
        //    Ds = new DataSet();
        //    Sda.Fill(Ds, "tabresult");
        //    DataTable SQLQuery = new DataTable();
        //    SQLQuery.Clear();
        //    SQLQuery = Ds.Tables["tabresult"];
        //    return SQLQuery;
        //}

        ////public static void SQLNonQuery(string Sql, Hashtable ParamTable = null)
        //{
        //    MySqlCommand CmdData = new MySqlCommand();
        //    // Warning!!! Optional parameters not supported
        //    Utility.Connect(); ();
        //    CmdData.CommandText = Sql;
        //    CmdData.Connection = con;
        //    CmdData.Transaction = Trans;
        //    CmdData.CommandTimeout = 180;
        //    if (ParamTable != null)
        //    {
        //        foreach (DictionaryEntry DeData in ParamTable)
        //        {
        //            CmdData.Parameters.AddWithValue(DeData.Key.ToString(), DeData.Value);
        //        }
        //    }
        //    CmdData.ExecuteNonQuery();
        //}

        //public static void SQLNonQuery1(string Sql, DataTable Dt, Hashtable ParamTable = null)
        //{
        //    try
        //    {
        //        MySqlCommand CmdData = new MySqlCommand();
        //        // Warning!!! Optional parameters not supported
        //        Utility.Connect();
        //        CmdData.CommandText = Sql;
        //        CmdData.Connection = Cn;
        //        CmdData.CommandType = CommandType.StoredProcedure;
        //        CmdData.Transaction = Trans;
        //        CmdData.CommandTimeout = 180;
        //        if (!(ParamTable == null))
        //        {
        //            foreach (DictionaryEntry DeData in ParamTable)
        //            {
        //                CmdData.Parameters.AddWithValue(DeData.Key.ToString(), DeData.Value);
        //            }

        //        }
        //        CmdData.Parameters.AddWithValue("@Dt", Dt);
        //        CmdData.ExecuteNonQuery();
        //    }
        //    catch (Exception ex) { MessageBox.Show(ex.Message.ToString()); }

        //}

        //public static object SQLScalar(string Sql, Hashtable ParamTable = null)
        //{
        //    MySqlCommand CmdData = new MySqlCommand();
        //    // Warning!!! Optional parameters not supported
        //    Utility.Connect();
        //    CmdData.CommandText = Sql;
        //    CmdData.Connection = con;
        //    CmdData.Transaction = Trans;
        //    if (ParamTable != null)
        //    {
        //        foreach (DictionaryEntry DeData in ParamTable)
        //        {
        //            CmdData.Parameters.AddWithValue(DeData.Key.ToString(), DeData.Value);
        //        }

        //    }

        //    return CmdData.ExecuteScalar();
        //}

      

        //public static void MySQL_Backup(string Path)
        //{
        //    string constring = GlobalVariables.ConnectionString;
        //    string file = Path;
        //    using (MySqlConnection conn = new MySqlConnection(constring))
        //    {
        //        using (MySqlCommand cmd = new MySqlCommand())
        //        {
        //            using (MySqlBackup mb = new MySqlBackup(cmd))
        //            {
        //                cmd.Connection = conn;
        //                conn.Open();
        //                mb.ExportToFile(file);
        //                conn.Close();
        //            }
        //        }
        //    }
        //}

        //public static void MySQL_Restore(string Path)
        //{
        //    string constring = GlobalVariables.ConnectionString;
        //    string file = Path;
        //    using (MySqlConnection conn = new MySqlConnection(constring))
        //    {
        //        using (MySqlCommand cmd = new MySqlCommand())
        //        {
        //            using (MySqlBackup mb = new MySqlBackup(cmd))
        //            {
        //                cmd.Connection = conn;
        //                conn.Open();
        //                mb.ImportFromFile(file);
        //                conn.Close();
        //            }
        //        }
        //    }
        //}
    }
    abstract class animal
    {
        public abstract void eat();
        public void sound()
        {
            Console.WriteLine("dog can sound");
        }
    }
    class dog : animal
    {
        public override void eat()
        {
            Console.WriteLine("dog can eat");
        }
    }
}
using System.Data.SqlClient;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace Pinnacle.Models
{
    public static class MsSQLAcc
    {
        public static SqlConnection Connection { get; set; }
        public static SqlTransaction Transaction { get; set; }
        public static void OpenConnection()
        {
            if (string.IsNullOrEmpty(GlobalVariables.MMISConnectionString))
            {
                throw new Exception("Connection string is not valid");
            }
            if (Connection == null)
            {
                Connection = new SqlConnection(GlobalVariables.MMISConnectionString);
            }
            if (Connection != null)
            {
                if (Connection.State == ConnectionState.Closed || Connection.State == ConnectionState.Broken)
                {
                    Connection.Open();
                }
            }
        }
        public static void CloseConnection()
        {
            if (Connection != null)
            {
                if (Connection.State == ConnectionState.Open)
                {
                    Connection.Close();
                }
            }
        }
        public static DataTable ExecuteQuery(string Query, Hashtable ParamTable = null)
        {
            OpenConnection();
            SqlCommand Cmd = new SqlCommand(Query, Connection, Transaction);
            Cmd.CommandTimeout = 180;
            if (ParamTable != null)
            {
                foreach (DictionaryEntry DeData in ParamTable)
                {
                    Cmd.Parameters.AddWithValue(DeData.Key.ToString(), DeData.Value);
                }
            }
            SqlDataAdapter Adp = new SqlDataAdapter(Cmd);
            var datatable = new DataTable();
            Adp.Fill(datatable);
            return datatable;
        }
        public static DataTable ExecuteQuerySP(string SPame, Hashtable ParamTbl = null)
        {
            OpenConnection();
            SqlCommand Cmd = new SqlCommand(SPame, Connection, Transaction);
            Cmd.CommandTimeout = 180;
            Cmd.CommandType = CommandType.StoredProcedure;
            if (ParamTbl != null)
            {
                foreach (DictionaryEntry DeData in ParamTbl)
                {
                    Cmd.Parameters.AddWithValue(DeData.Key.ToString(), DeData.Value);
                }
            }
            SqlDataAdapter Adp = new SqlDataAdapter(Cmd);
            var datatable = new DataTable();
            Adp.Fill(datatable);
            return datatable;
        }
        public static void ExecuteNonQuery(string Query, Hashtable ParamTbl = null)
        {
            OpenConnection();
            SqlCommand Cmd = new SqlCommand(Query, Connection, Transaction);
            Cmd.CommandTimeout = 180;
            Cmd.CommandType = CommandType.Text;
            if (ParamTbl != null)
            {
                foreach (DictionaryEntry DeData in ParamTbl)
                {
                    Cmd.Parameters.AddWithValue(DeData.Key.ToString(), DeData.Value);
                }
            }
            Cmd.ExecuteNonQuery();
        }
        public static void ExecuteNonQuerySP(string SPName, Hashtable ParamTbl = null)
        {
            OpenConnection();
            SqlCommand Cmd = new SqlCommand(SPName, Connection, Transaction);
            Cmd.CommandTimeout = 180;
            Cmd.CommandType = CommandType.StoredProcedure;
            if (ParamTbl != null)
            {
                foreach (DictionaryEntry DeData in ParamTbl)
                {
                    Cmd.Parameters.AddWithValue(DeData.Key.ToString(), DeData.Value);
                }
            }
            Cmd.ExecuteNonQuery();
        }
        public static object ExecuteScalar(string Query, Hashtable ParamTable = null)
        {
            OpenConnection();
            SqlCommand Cmd = new SqlCommand(Query, Connection);
            Cmd.CommandType = CommandType.Text;
            if (ParamTable != null)
            {
                foreach (DictionaryEntry DeData in ParamTable)
                {
                    Cmd.Parameters.AddWithValue(DeData.Key.ToString(), DeData.Value);
                }
            }
            object obj = Cmd.ExecuteScalar();
            return obj;
        }
        //public static List<T> ExecuteReader<T>(string Query, Hashtable ParamTable = null)
        //{
        //    OpenConnection();
        //    SqlCommand Cmd = new SqlCommand(Query, Connection, Transaction);
        //    Cmd.CommandType = CommandType.Text;
        //    if (ParamTable != null)
        //    {
        //        foreach (DictionaryEntry DeData in ParamTable)
        //        {
        //            Cmd.Parameters.AddWithValue(DeData.Key.ToString(), DeData.Value);
        //        }
        //    }
        //    List<T> list = new List<T>();
        //    IDataReader reader = Cmd.ExecuteReader(CommandBehavior.Default);
        //    while (reader.Read())
        //    {
        //        var row = Activator.CreateInstance<T>();
        //        for (int i = 0; i <= reader.FieldCount - 1; i++)
        //        {
        //            var property = row.GetType().GetProperty(reader.GetName(i));
        //            if (property != null && !object.Equals(reader[i], DBNull.Value))
        //            {
        //                if (reader[property.Name].GetType() == typeof(Guid))
        //                {
        //                    property.SetValue(row, reader[property.Name].ToString());
        //                }
        //                else
        //                {
        //                    property.SetValue(row, Convert.ChangeType(reader[property.Name], property.PropertyType));
        //                }
        //            }
        //        }
        //        list.Add(row);
        //    }
        //    reader.Close();
        //    return list;
        //}
       //public static List<T> ExecuteReaderSP<T>(string SPName, Hashtable ParamTable = null)
       // {
       //     OpenConnection();
       //     SqlCommand Cmd = new SqlCommand(SPName, Connection, Transaction);
       //     Cmd.CommandType = CommandType.StoredProcedure;
       //     if (ParamTable != null)
       //     {
       //         foreach (DictionaryEntry DeData in ParamTable)
       //         {
       //             Cmd.Parameters.AddWithValue(DeData.Key.ToString(), DeData.Value);
       //         }
       //     }
       //     List<T> list = new List<T>();
       //     IDataReader reader = Cmd.ExecuteReader(CommandBehavior.Default);
       //     while (reader.Read())
       //     {
       //         var row = Activator.CreateInstance<T>();
       //         for (int i = 0; i <= reader.FieldCount - 1; i++)
       //         {
       //             var property = row.GetType().GetProperty(reader.GetName(i));
       //             if (property != null && !object.Equals(reader[i], DBNull.Value))
       //             {
       //                 if (property.PropertyType == typeof(Guid))
       //                 {
       //                     property.SetValue(row, reader[property.Name].ToString());
       //                 }
       //                 else
       //                 {
       //                     property.SetValue(row, Convert.ChangeType(reader[property.Name], property.PropertyType));
       //                 };
       //             }
       //         }
       //         list.Add(row);
       //     }
       //     reader.Close();
       //     return list;
       // }
    }
}

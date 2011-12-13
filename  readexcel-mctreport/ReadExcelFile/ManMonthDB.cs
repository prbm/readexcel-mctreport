using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Npgsql;
using System.Windows.Forms;

namespace ReadExcelFile
{
    class ManMonthDB
    {
        private String server;
        private String database;
        private String port;
        private String user;
        private String pwd;
        private String connectionString = null;
        private NpgsqlConnection conn = null;

        public ManMonthDB()
        {
            server = "127.0.0.1";
            port = "5432";
            user = "postgres";
            pwd = "postgres";
            database = "ManMonth";

            connectionString = "Server=" + server + ";Port=" + port + ";User Id=" + user +
                               ";Password=" + pwd + ";Database=" + database + ";";

            conn = new NpgsqlConnection(connectionString);
        }

        public bool openConnection()
        {
            try
            {
                if (conn != null)
                {
                    conn.Open();
                    MessageBox.Show("Conectou");
                    return true;
                }
                else
                    return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
        }

        public bool closeConnection()
        {
            try
            {
                if (conn != null)
                {
                    conn.Close();
                    MessageBox.Show("Desconectou");
                    return true;
                }
                else
                    return false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
        }

        public List<Object> select(String sqlCmd, Type t)
        {
            List<Object> result = new List<Object>();
            NpgsqlCommand sql = new NpgsqlCommand(sqlCmd);
            sql.Connection = conn;
            NpgsqlDataReader dr = sql.ExecuteReader();

            while (dr.Read())
            {
                    Object[] tmp = new Object[dr.FieldCount];
                    if (t.Name.ToString().Equals("PMSStatus"))
                    {
                        dr.GetValues(tmp);
                        PMSStatus pmsSt = new PMSStatus();
                        String name = tmp[0].GetType().Name;
                        pmsSt.ID = Convert.ToInt32(tmp[0]);
                        pmsSt.Code = ((String)tmp[1]).Trim();
                        pmsSt.Description = ((String)tmp[2]).Trim();
                        pmsSt.CreationDate = (DateTime)tmp[3];

                        result.Add(pmsSt);
                    }

                    else if (t.Name.ToString().Equals("Subsidiary"))
                    {
                        dr.GetValues(tmp);
                        Subsidiary sub = new Subsidiary();
                        String name = tmp[0].GetType().Name;
                        sub.ID = Convert.ToInt32(tmp[0]);
                        sub.Name = ((String)tmp[1]).Trim();
                        sub.Description = ((String)tmp[2]).Trim();

                        result.Add(sub);
                    }
            }

            return result;
        }

        public bool insert(String cmd)
        {
            NpgsqlCommand sqlCmd = new NpgsqlCommand(cmd);

            if (conn != null)
            {
                sqlCmd.Connection = conn;
                Int32 numLinhas = sqlCmd.ExecuteNonQuery();
                if (numLinhas > 0)
                    return true;
                else
                return false;
            }

            return false;
        }

        public Int32 countRegistersOnTable(String tableName)
        {
            Int32 result = 0;
            NpgsqlCommand sqlCmd = new NpgsqlCommand("SELECT count(*) FROM " + tableName);

            if (conn != null)
            {
                sqlCmd.Connection = conn;
                NpgsqlDataReader dr = sqlCmd.ExecuteReader();

                while (dr.Read())
                    result = Int32.Parse(dr.GetValue(0).ToString());
            }

            return result;
        }

        public Int32 selectMaxID(String tableName)
        {
            Int32 result = 0;
            NpgsqlCommand sqlCmd = new NpgsqlCommand("SELECT MAX(id) FROM " + tableName);

            if (conn != null)
            {
                sqlCmd.Connection = conn;
                NpgsqlDataReader dr = sqlCmd.ExecuteReader();

                while (dr.Read())
                    result = Int32.Parse(dr.GetValue(0).ToString());
            }

            return result;
        }
    }
}

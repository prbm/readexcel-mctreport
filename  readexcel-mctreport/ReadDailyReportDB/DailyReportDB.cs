using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MySql.Data.MySqlClient;
using System.Windows;


namespace ReadDailyReportDB
{
    class DailyReportDB
    {
        private String server;
        private String port;
        private String database;
        private String user;
        private String pwd;
        private String connectionString = null;
        private MySqlConnection conn = null;

        public DailyReportDB()
        {
            server = "10.193.225.22";
            port = "3306";
            database = "dailyreport";
            user = "readonly";
            pwd = "ylnodaer";
            connectionString = "SERVER=" + server + ";" + "DATABASE=" + database + ";";
            connectionString += "UID=" + user + ";" + "PASSWORD=" + pwd + ";";
            conn = new MySqlConnection(connectionString);
        }

        public bool openConnection()
        {
            try
            {
                conn.Open();
                return true;
            }
            catch (MySqlException ex)
            {
                switch (ex.Number)
                {
                    case 0:
                        MessageBox.Show("Can not connect to database.", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        break;
                    case 1045:
                        MessageBox.Show("Invalid User Name/Password", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        break;
                }

                return false;
            }
        }

        public bool closeConnection()
        {
            try
            {
                conn.Close();
                return true;
            }
            catch (MySqlException ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
        }

        public List<Object> executeSelect(String query, String res)
        {
            List<Object> list = new List<Object>();
            MySqlCommand cmd = new MySqlCommand(query, conn);

            MySqlDataReader dataReader = cmd.ExecuteReader();

            while (dataReader.Read())
            {
                list.Add(dataReader[res]);
            }

            dataReader.Close();

            return list;
        }
    }
}

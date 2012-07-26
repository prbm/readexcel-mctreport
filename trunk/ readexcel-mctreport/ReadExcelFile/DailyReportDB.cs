using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MySql.Data.MySqlClient;

namespace ReadExcelFile
{
    class DailyReportDB
    {
        private String server;
        private String database;
        private String port;
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

        public DailyReportDB(String database)
        {
            server = "10.193.225.22";
            port = "3306";
            this.database = database;
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
                if (conn != null)
                {
                    conn.Open();
                    //MessageBox.Show("Conectou");
                    return true;
                }
                else
                    return false;
            }
            catch (MySqlException ex)
            {
                switch (ex.Number)
                {
                    case 0:
                        MessageBox.Show("Can not connect to database.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        break;
                    case 1045:
                        MessageBox.Show("Invalid User Name/Password", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        break;
                }

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
                    //MessageBox.Show("Desconectou");
                    return true;
                }
                else
                    return false;
            }
            catch (MySqlException ex)
            {
                MessageBox.Show(ex.Message);
                return false;
            }
        }

        public Int32 selectCount(String sqlCmd)
        {
            MySqlCommand sql = new MySqlCommand(sqlCmd);

            sql.Connection = conn;
            Int32 result = 0;
            MySqlDataReader dr = sql.ExecuteReader();
            while (dr.Read())
            {
                result = Convert.ToInt32(dr.GetValue(0));
            }

            return result;
        }

        public List<Object> select(String sqlCmd, Type t)
        {
            List<Object> result = new List<Object>();
            MySqlCommand sql = new MySqlCommand(sqlCmd);
            sql.Connection = conn;
            MySqlDataReader dr = sql.ExecuteReader();

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

                else if (t.Name.ToString().Equals("Employee"))
                {
                    dr.GetValues(tmp);
                    Employee emp = new Employee();
                    String name = tmp[0].GetType().Name;
                    emp.Id = Convert.ToInt32(tmp[0]);
                    emp.Name = ((String)tmp[3]).Trim();

                    result.Add(emp);
                }
                else if (t.Name.ToString().Equals("Model"))
                {
                    Model m = new Model();
                    m.ModelID = Convert.ToInt32(dr.GetValue(0));
                    m.ModelCode = ((String)dr.GetValue(1)).Trim();

                    result.Add(m);
                }
                else if (t.Name.ToString().Equals("ProjectCourseCarrier"))
                {
                    ProjectCourseCarrier m = new ProjectCourseCarrier();
                    m.IdCarrier = Convert.ToInt32(dr.GetValue(0));
                    m.IdCountry = Convert.ToInt32(dr.GetValue(1));
                    m.Name = ((String)dr.GetValue(2)).Trim();

                    result.Add(m);
                }
                else if (t.Name.ToString().Equals("ProjectReportedHours"))
                {
                    ProjectReportedHours m = new ProjectReportedHours();
                    m.ProjectId = Convert.ToInt32(dr.GetValue(1));
                    m.AmountReported = ((TimeSpan)dr.GetValue(3)).TotalHours;
                    m.EmployeeId = Convert.ToInt32(dr.GetValue(13));
                    m.Month = ((DateTime)dr.GetValue(14)).Month;
                    m.Year = ((DateTime)dr.GetValue(14)).Year;

                    result.Add(m);
                }
                else
                {
                    tmp = new Object[dr.FieldCount];
                    dr.GetValues(tmp);

                    result.Add(tmp);
                }
            }

            return result;
        }

        public bool insert(String cmd)
        {
            MySqlCommand sqlCmd = new MySqlCommand(cmd);

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
    }
}

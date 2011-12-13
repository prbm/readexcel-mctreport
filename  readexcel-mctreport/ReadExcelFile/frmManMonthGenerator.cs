using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ReadExcelFile
{
    public partial class frmManMonthGenerator : Form
    {
        PMSStatusCollection pmsStColl = new PMSStatusCollection();
        List<PMSStatus> stColl = new List<PMSStatus>();
        List<Subsidiary> subsidiaries = new List<Subsidiary>();

        public frmManMonthGenerator()
        {
            InitializeComponent();
        }

        private void pbBrowsePMSWorkbook_Click(object sender, EventArgs e)
        {
            // Declara variable to store the selected file created through PMS
            OpenFileDialog oFD = new OpenFileDialog();
            oFD.InitialDirectory = "C:\\";
            oFD.Filter = "Excel files 1997-2003 (*.xls)|*.xls|Excel files >2007 (*.xlsx)|*.xlsx";
            oFD.FilterIndex = 1;

            try
            {
                // check if a new file was chosen
                if (oFD.ShowDialog() == DialogResult.OK)
                    edPMSWorkbook.Text = oFD.FileName.Trim();
                else
                    return;

            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
            finally
            {
                Cursor.Current = Cursors.Default;
            }
        }

        private void pbGetPMSData_Click(object sender, EventArgs e)
        {

            String[,] values;
            ExcelRead er = new ExcelRead();
            DateTime today = DateTime.Now;

            Cursor.Current = Cursors.WaitCursor;

            /***************************************
             * * COLLECT DATA FROM PMS WOROKBOOKS **
             * ************************************/

            // Define the indexes of the select WorkSheets
            Object[] selectWorksheetsIndexes = { 0 };

            // Collect the indexes of the select columns
            Object[] selectColumnsIndexes = {0, 1, 2, 4, 14};

            er.processWorkSheets(edPMSWorkbook.Text, selectWorksheetsIndexes, selectColumnsIndexes);
            values = new String[er.CellValues.GetLength(0), er.CellValues.GetLength(1)];
            values = er.CellValues;

            Int32 contLinhas = er.CellValues.GetLength(0);
            List<ProjectCA> pCAs = new List<ProjectCA>();

            er = null;

            // collect PMS status stored in database
            ManMonthDB mmDB = new ManMonthDB();
            mmDB.openConnection();
            List<Object> obj = mmDB.select("SELECT * FROM pms_status_desc", typeof(PMSStatus));
            foreach (PMSStatus o in obj)
            {
                pmsStColl.Add(o);
                stColl.Add(o);
            }

            obj = mmDB.select("SELECT * FROM subsidiary", typeof(Subsidiary));
            foreach (Subsidiary o in obj)
                subsidiaries.Add(o);

            ////MessageBox.Show("Number of registers on table: " + mmDB.countRegistersOnTable("project_ca"));
            //MessageBox.Show("Number of registers on table: " + pmsStColl.Count);
            mmDB.closeConnection();

            // check in database if there are registers to be counted.
            String tmpText = null;
            Country country = new Country();
            Carrier carrier = new Carrier();

            for (int cont = 0; cont < contLinhas; cont++)
            {
                ProjectCA pCA = new ProjectCA();
                pCA.ProjectCode = values[cont, 0];
              
                // normalize the country and carrier names
                country.Name = values[cont, 1].Trim().ToUpper();
                if (country.Name.Equals("NO COUNTRY NAME"))
                {
                    if (pCA.ProjectCode.Contains(".AVEN"))
                    {
                        pCA.CountryName = "VENEZUELA";
                        carrier.Name = "OPEN";
                    }
                    else if (pCA.ProjectCode.Contains(".APRY"))
                    {
                        pCA.CountryName = "PARAGUAY";
                        carrier.Name = "OPEN";
                    }
                    else
                    {
                        pCA.CountryName = "BASIS";
                        carrier.Name = "OPEN";
                    }

                }
                else
                {
                    pCA.CountryName = country.Name;
                    carrier.Name = values[cont, 2];
                    pCA.CarrierName = carrier.Name;
                }


                // get status code
                tmpText = values[cont, 4];
                if (tmpText != null)
                {
                    PMSStatus st = stColl.Find(delegate(PMSStatus pSt) {return (pSt.Description.ToUpper().Trim().Equals(tmpText.ToUpper().Trim()));});
                    if(st!=null)
                        pCA.PmsStatus= st;
                    else
                        pCA.PmsStatus = stColl.Find(delegate(PMSStatus pSt) { return (pSt.ID == 0); });
                }
               
                // all handset development is taken as a R&D Project
                pCA.PdStatusProject = 1;

                pCA.PcaStatus.Nyear = today.Year;
                pCA.PcaStatus.Nmonth = today.Month;
                pCA.PcaStatus.Status = pCA.PmsStatus.ID;
                pCA.PcaStatus.DateOfChange = today;
                pCA.PcaStatus.ProjectID = 0;

                pCAs.Add(pCA);
            }





            /**********************************
             * * READ USER IDS FROM DATABASE **
             * ********************************/
            List<Employee> emps = new List<Employee>();
            DailyReportDB drDb = new DailyReportDB("weekly_innodb");
            if (!drDb.openConnection())
                return;

            String dailyIdUsers = "(";

            obj = drDb.select("SELECT * FROM user WHERE user.idTeam=1", typeof(Employee));
            foreach (Employee employee in obj)
            {
                emps.Add(employee);
                dailyIdUsers += "daily.idUser = " + employee.Id + " || ";
            }

            dailyIdUsers = dailyIdUsers.Substring(0, dailyIdUsers.Length - 4) + ")";

            MessageBox.Show("Insert: " + dailyIdUsers);

            if (!drDb.closeConnection())
                return;


            String cmd = "SELECT SUM(TIME_TO_SEC(effortxproject.time)) FROM effortxproject INNER JOIN (effort, daily) ";
            cmd += "ON (daily.idDaily=effort.idDaily AND effortxproject.idEffort = effort.idEffort) WHERE ";
            cmd += "(daily.date > '2011-10-31' AND daily.date<'2011-12-01') AND ";
            cmd += "(effortxproject.idModel = 190) AND " + dailyIdUsers;

            drDb = new DailyReportDB();

            if (!drDb.openConnection())
                return;

            MessageBox.Show("Calculated Hours = " + (Convert.ToDouble(drDb.selectCount(cmd)) / 3600).ToString("#.##"));

            if (!drDb.closeConnection())
                return;

            return;







            // store project data in database
            
            mmDB.openConnection();

            foreach (ProjectCA p in pCAs)
            {
                Subsidiary s = new Subsidiary();
                s = subsidiaries.Find(delegate (Subsidiary ss) { return (ss.Name==s.getSubsidiaryCode(p.CountryName));});

                String insertCmd1 = "INSERT INTO project_ca (projectcode, carriername, countryname, rdstatusproject, subsidiaryid) values ";
                String insertCmd2 = "INSERT INTO project_status (nyear, nmonth, status, dateofchange, projectid) values ";

                insertCmd1 += "('" + p.ProjectCode + "','" + p.CarrierName + "','" + p.CountryName + "'," + p.PdStatusProject + ",";
                if(s!=null)
                    insertCmd1 += s.ID + ")";
                else
                    insertCmd1 += "0)";

                insertCmd2 += "(" + p.PcaStatus.Nyear + "," + p.PcaStatus.Nmonth + "," + p.PcaStatus.Status + ",'";
                insertCmd2 += p.PcaStatus.DateOfChange.ToString("yyyy-MM-dd") + "'," ;

                if(mmDB.insert(insertCmd1))
                {
                    Int32 id = mmDB.selectMaxID("project_ca");
                    insertCmd2 += id + ")";
                    mmDB.insert(insertCmd2);
                }

            }

            mmDB.closeConnection();
            Cursor.Current = Cursors.Default;

            MessageBox.Show("FINI");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ManMonthDB mmDB = new ManMonthDB();
            mmDB.openConnection();
            List<Object> obj = mmDB.select("SELECT * FROM pms_status_desc", typeof(PMSStatus));
            foreach (PMSStatus o in obj)
                pmsStColl.Add(o);

            //MessageBox.Show("Number of registers on table: " + mmDB.countRegistersOnTable("project_ca"));
            MessageBox.Show("Number of registers on table: " + pmsStColl.Count);
            mmDB.closeConnection();

        }

        private void pbMySQL_Click(object sender, EventArgs e)
        {
            DailyReportDB drDB = new DailyReportDB();
            drDB.openConnection();

            String cmd = "SELECT SUM(TIME_TO_SEC(effortxproject.time)) FROM effortxproject INNER JOIN (effort, daily) ";
            cmd += "ON (daily.idDaily=effort.idDaily AND effortxproject.idEffort = effort.idEffort) WHERE ";
            cmd += "(daily.date > '2011-10-31' AND daily.date<'2011-12-01') AND ";
            cmd += "(effortxproject.idModel = 190) AND (daily.idUser = 138 || daily.idUser = 1)";


            MessageBox.Show("Calculated Hours = " + (Convert.ToDouble(drDB.selectCount(cmd)) / 3600).ToString("#.##"));

            drDB.closeConnection();


        }
    }
}

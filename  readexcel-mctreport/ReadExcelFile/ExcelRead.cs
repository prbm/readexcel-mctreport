using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Globalization;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;

namespace ReadExcelFile
{
    class ExcelRead
    {
        Excel.Application excelApp;
        List<ProjectDB> projectList;
        private List<String> columnHeads;
        private List<String> wSheets;
        private String[,] cellValues;
        int linha = 0, planilha = 0;

        public ExcelRead()
        {
            excelApp = new Excel.Application();
            this.projectList = new List<ProjectDB>();
            cellValues = null;
        }

        public void getWorkSheets(String filePath)
        {
            Excel._Workbook excelWBook;
            Excel._Worksheet excelWSheet;
            wSheets = new List<String>();

            try
            {
                excelWBook = excelApp.Workbooks.Open(filePath, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, false, Type.Missing, Type.Missing, false, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                for (int cWSheet = 1; cWSheet <= excelWBook.Worksheets.Count; cWSheet++)
                {
                    excelWSheet = (Excel.Worksheet)excelWBook.Worksheets.get_Item(cWSheet);
                    wSheets.Add(excelWSheet.Name.ToString());
                }

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            finally
            {
                // close Excel Application
                excelApp.Quit();
            }
        }

        // get data related to all columns in all selected worksheets
        public void processWorkSheets(String filePath, Object[] selectedWorkSheets, Object[] selectedColumnIndexes)
        {
            Excel._Workbook excelWBook;
            Excel._Worksheet excelWSheet;
            columnHeads = new List<String>();
            Object obj = null;

            // get the emnumerator of elements in each array
            System.Collections.IEnumerator numberSelectedWorkSheetProcess = selectedWorkSheets.GetEnumerator();
            System.Collections.IEnumerator numberSelectedColumnHeads = selectedColumnIndexes.GetEnumerator();

            Int32 rowCount = 0, columnCount = 0;
            String[,] cells = null;

            String culture = Thread.CurrentThread.CurrentCulture.Name;

            //if (!culture.Equals("en-US"))
            //    Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            String format = "yyyy/MM/dd";

            try
            {
                excelWBook = excelApp.Workbooks.Open(filePath, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, false, Type.Missing, Type.Missing, false, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                // get the total number of lines in all selected WorkSheets
                while (numberSelectedWorkSheetProcess.MoveNext())
                {
                    excelWSheet = (Excel.Worksheet)excelWBook.Worksheets.get_Item((Int32)numberSelectedWorkSheetProcess.Current + 1);
                    rowCount += (excelWSheet.UsedRange.Rows.Count - 1);
                }

                numberSelectedWorkSheetProcess.Reset();

                // read all sreapsheets existent in the Excel File
                while (numberSelectedWorkSheetProcess.MoveNext())
                {
                    // get the content of the current SpreadSheet
                    excelWSheet = (Excel.Worksheet)excelWBook.Worksheets.get_Item((Int32)numberSelectedWorkSheetProcess.Current + 1);

                    // get the range of used columns and rows
                    Excel.Range cellsRange = excelWSheet.UsedRange;

                    if (cells == null)
                    {
                        cells = new String[rowCount, selectedColumnIndexes.Length];
                        rowCount = 0;
                    }

                    // read all lines in the working spreadsheet
                    for (int rCount = 2; rCount <= cellsRange.Rows.Count; rCount++)
                    {
                        columnCount = 0;
                        // read selected columns
                        while (numberSelectedColumnHeads.MoveNext())
                        {
                            Int32 selectedColumn = ((Int32)numberSelectedColumnHeads.Current + 1);
                            // get data from the seleceted column
                            if ((cellsRange.Cells[rCount, selectedColumn] as Excel.Range).Value2 != null)
                            {
                                obj = (cellsRange.Cells[rCount, selectedColumn] as Excel.Range).Value2;
                                String numberFormat = (cellsRange.Cells[rCount, selectedColumn] as Excel.Range).NumberFormat;

                                if (obj.GetType() == typeof(double))
                                {
                                    if (numberFormat.ToLower().Contains("yyyy"))
                                    {
                                        DateTime date = DateTime.FromOADate(Convert.ToDouble(obj));
                                        String tmp = null;
                                        if(Thread.CurrentThread.CurrentCulture.Name.Equals("en-US"))
                                            tmp = date.ToString("MM/dd/yy");
                                        else
                                            tmp = date.ToString("dd/MM/yy");
                                        //String tmp = date.ToString(dtTimeFmt.ShortDatePattern);
                                        String[] dt = tmp.Split('/');
                                        dt[0] = "20" + dt[0];
                                        DateTime dtime = new DateTime(Int32.Parse(dt[0]), Int32.Parse(dt[1]), Int32.Parse(dt[2]));
                                        cells[rowCount, columnCount] = dtime.ToString(format);
                                    }
                                    else
                                        cells[rowCount, columnCount] = Double.Parse(obj.ToString()).ToString();

                                }
                                else if (obj.ToString().ToUpper().Contains("(C)") || obj.ToString().ToUpper().Contains("(D)") || obj.ToString().ToUpper().Contains("(H)"))
                                {
                                    String tmp = obj.ToString();
                                    tmp = tmp.Substring(tmp.IndexOf('(') - 8, 8);

                                    String[] date = tmp.Split('/');
                                    date[0] = "20" + date[0];
                                    DateTime dt;

                                    if (Thread.CurrentThread.CurrentCulture.Name.Equals("en-US"))
                                        dt = new DateTime(Int32.Parse(date[0]), Int32.Parse(date[2]), Int32.Parse(date[1]));
                                    else
                                        dt = new DateTime(Int32.Parse(date[0]), Int32.Parse(date[1]), Int32.Parse(date[2]));

                                    cells[rowCount, columnCount] = dt.ToString(format);
                                }
                                else if (obj.ToString().Contains(@"→"))
                                {
                                    String tmp = obj.ToString();

                                    while (tmp.Contains('→'))
                                        tmp = tmp.Substring(tmp.IndexOf('→') + 1);

                                    String[] date = tmp.Split('/');
                                    date[0] = "20" + date[0];
                                    DateTime dt;

                                    if (Thread.CurrentThread.CurrentCulture.Name.Equals("en-US"))
                                        dt = new DateTime(Int32.Parse(date[0]), Int32.Parse(date[2]), Int32.Parse(date[1]));
                                    else
                                        dt = new DateTime(Int32.Parse(date[0]), Int32.Parse(date[1]), Int32.Parse(date[2]));

                                    cells[rowCount, columnCount] = dt.ToString(format);
                                }
                                else
                                    cells[rowCount, columnCount] = obj.ToString(); // fix the date format from PMS to Brazilian format
                            }
                            else
                                cells[rowCount, columnCount] = "";

                            columnCount++;
                        }

                        // return to the first item of the array
                        numberSelectedColumnHeads.Reset();
                        rowCount++;
                    }
                }
                excelWBook.Close(true, null, null);
                cellValues = cells;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            finally
            {
                // close Excel Application
                excelApp.Quit();
                Thread.CurrentThread.CurrentCulture = new CultureInfo(culture);
            }
        }

        public void calculateManMonth(String filePath, String fileDestPath, String fileProjectStatusPath, Object[] selectedWorkSheets, Object[] selectedColumnIndexes, List<CountryCode> listCountryCodes)
        {
            Excel._Workbook excelWBook;
            Excel._Worksheet excelWSheet;
            Object obj = null;
            Int32 columnCount = 0;
            String tmp = null;
            Model m;
            CA ca;
            List<Model> modelList = new List<Model>();
            String msg = @"INI: " + DateTime.Now.ToString("hh:mm:ss");

            // get all selected columns
            columnCount = 0;
            int[] selectedColumns = new int[selectedColumnIndexes.Length];

            foreach (int a in selectedColumnIndexes)
                selectedColumns[columnCount++] = a;

            // get the emnumerator of elements in each array
            System.Collections.IEnumerator numberSelectedWorkSheetProcess = selectedWorkSheets.GetEnumerator();

            try
            {
                // open the spreadsheet with the source data
                excelWBook = excelApp.Workbooks.Open(filePath, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                // move to the first spreadsheet in the document
                numberSelectedWorkSheetProcess.Reset();

                // read all sreapsheets existent in the Excel File
                while (numberSelectedWorkSheetProcess.MoveNext())
                {
                    // get the content of the current SpreadSheet
                    excelWSheet = (Excel.Worksheet)excelWBook.Worksheets.get_Item((Int32)numberSelectedWorkSheetProcess.Current + 1);
                    planilha = (Int32)numberSelectedWorkSheetProcess.Current + 1;

                    // get the range of used columns and rows
                    Excel.Range cellsRange = excelWSheet.UsedRange;
                    ExcelColumns ec = new ExcelColumns(cellsRange.Columns.Count);
                    String[] head = ec.Columns;
                    String cell = null;

                    // read all lines in the working spreadsheet
                    for (int rCount = 2; rCount < cellsRange.Rows.Count; rCount++)
                    {
                        linha = rCount;

                        m = new Model();
                        ca = new CA(listCountryCodes);

                        // read model name
                        cell = head[selectedColumns[0]].ToString() + rCount.ToString();
                        tmp = (string)cellsRange.get_Range(cell, cell).Value2;
                        m.ModelCode = tmp.ToUpper().Trim();

                        // read country name
                        cell = head[selectedColumns[1]].ToString() + rCount.ToString();
                        tmp = (string)cellsRange.get_Range(cell, cell).Value2;
                        ca.Country = tmp.ToUpper().Trim();

                        // check the subsidiary name for the selected country
                        ca.setSubsidiary(ca.Country);
                        
                        // read carrier name
                        cell = head[selectedColumns[2]].ToString() + rCount.ToString();
                        tmp = (string)cellsRange.get_Range(cell, cell).Value2;
                        tmp = tmp.ToUpper().Trim();
                        //if (ca.Country.Equals("ARGENTINA"))
                        if (ca.Country.Equals("UNIFIED"))
                        {
                            int a = 999;
                        }
                        if (ca.Country.Equals("ARGENTINA") && tmp.Equals("CLARO"))
                            ca.CarrierName = "CTI";
                        else
                        {
                            ca.CarrierName = tmp.ToUpper().Trim();
                            if(ca.CarrierName.Equals("TIGO") || ca.CarrierName.Equals("TELEFONICA") || ca.CarrierName.Equals("CLARO"))
                                if(ca.Country.Equals("UNIFIED"))
                                    ca.Country = "CENTRAL AMERICA";
                        }
                        ////else if (ca.Country.Equals("UNIFIED") && tmp.Equals("TIGO"))
                        //else if (ca.Country.Equals("UNIFIED") && (tmp.Equals("TIGO") || tmp.Equals("TELEFONICA") || tmp.Equals("CLARO")))
                        //{
                        //    ca.Country = "CENTRAL AMERICA";
                        //    ca.CarrierName = tmp;
                        //    }
                        //else
                            

                        // read man month value
                        ca.MediumManMonth = 0.0;
                        cell = head[selectedColumns[3]].ToString() + rCount.ToString();
                        obj = cellsRange.get_Range(cell, cell).Value2;
                        if (obj.GetType() == typeof(string))
                            ca.MediumManMonth = Double.Parse((string)obj);
                        else if (obj.GetType() == typeof(double))
                            ca.MediumManMonth = (double)obj;

                        //ca.MediumManMonth = ca.MediumManMonth / 1000;

                        // read the number of people that reported hours in the project
                        cell = head[selectedColumns[4]].ToString() + rCount.ToString();
                        obj = cellsRange.get_Range(cell, cell).Value2;
                        //ca.PeopleReportedHours = 0;
                        if (obj.GetType() == typeof(string))
                            ca.PeopleReportedHours = Int32.Parse((string)obj);
                        else if (obj.GetType() == typeof(double))
                            ca.PeopleReportedHours = Int32.Parse(obj.ToString());

                        Model tmpModel = modelList.Find(delegate(Model mm) { return (mm.ModelCode.ToUpper().Equals(m.ModelCode.ToUpper()) && mm.ModelCA.CarrierName.ToUpper().Equals(ca.CarrierName.ToUpper()) && mm.ModelCA.Country.ToUpper().Equals(ca.Country.ToUpper()) && mm.ModelCA.MediumManMonth > 0); });
                        // add/update model information in the list
                        if (tmpModel == null){
                            ca.setRepHour(ca.MediumManMonth, excelWSheet.Name);
                                m.ModelCA = ca;
                                modelList.Add(m);
                            }
                        else{
                            // if a project already exists, update the value of manmonth for the project
                            ca.ListReportedHours = tmpModel.ModelCA.ListReportedHours;

                            ReportHours rH = ca.ListReportedHours.Find(delegate(ReportHours rh) { return rh.TeamName.ToUpper().Equals(excelWSheet.Name.ToUpper()); });
                            if (rH != null)
                                ca.setRepHour((ca.MediumManMonth + rH.ReportedTime), excelWSheet.Name);
                            else
                                ca.setRepHour(ca.MediumManMonth, excelWSheet.Name);
                            
                            ca.MediumManMonth += tmpModel.ModelCA.MediumManMonth;
                            modelList.Remove(tmpModel);
                            tmpModel.ModelCA = ca;
                            modelList.Add(tmpModel);
                            tmpModel = null;
                        }
                    }
                }
                excelWBook.Close(true, null, null);

                /***********************************
                 * read the status of the projects *
                 * *********************************/

                // open the excel file
                excelWBook = excelApp.Workbooks.Open(fileProjectStatusPath, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, false, Type.Missing, Type.Missing, false, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                // get the content of the current SpreadSheet
                excelWSheet = (Excel.Worksheet)excelWBook.Worksheets.get_Item(1);
                // get the used ranges in the excel spreadsheet
                Excel.Range excelRange = excelWSheet.UsedRange;
                // read the content of the spreadSheet
                List<Model> listModelPjtStatus = new List<Model>();
                for (int counter = 2; counter <= excelRange.Rows.Count; counter++)
                {
                    m = new Model();
                    ca = new CA(listCountryCodes);

                    linha = counter;

                    if (linha == 316)
                        linha = linha;

                    // read model name
                    tmp = (string)excelRange.get_Range("A" + counter, "A" + counter).Value2;
                    m.ModelCode = tmp.ToUpper().Trim();
                    m.ProjectCode = tmp.ToUpper().Trim();

                    // read country name
                    tmp = (string)excelRange.get_Range("B" + counter, "B" + counter).Value2;
                    ca.Country = tmp.ToUpper().Trim();

                    // if the country name is not in the list, add it to the list
                    if (listCountryCodes.Find(delegate(CountryCode c) { return (c.Country == ca.Country); }) == null)
                        ca.Country = ca.Country.Trim();

                    // read carrier name
                    tmp = (string)excelRange.get_Range("C" + counter, "C" + counter).Value2;
                    ca.CarrierName = tmp.ToUpper().Trim();

                    // if the carrier name is not in the list, add it to the list
                    if (listCountryCodes.Find(delegate(CountryCode c) { return (c.Carrier == ca.CarrierName); }) == null)
                        ca.CarrierName = ca.CarrierName;

                    // read project status
                    tmp = (string)excelRange.get_Range("E" + counter, "E" + counter).Value2;
                    ca.ProjectStatus = tmp;

                    // update information of project status
                    if (modelList.Find(delegate(Model mm) { return (mm.ModelCode.Equals(m.ModelCode) && mm.ModelCA.CarrierName.Equals(ca.CarrierName) && mm.ModelCA.Country.Equals(ca.Country)); }) != null)
                    {
                        // get model data
                        Model mtmp = modelList.Find(delegate(Model mm) { return (mm.ModelCode.Equals(m.ModelCode) && mm.ModelCA.CarrierName.Equals(ca.CarrierName) && mm.ModelCA.Country.Equals(ca.Country)); });
                        // update model data with status information
                        modelList.Remove(mtmp);
                        mtmp.ModelCA.ProjectStatus = ca.ProjectStatus;
                        mtmp.ProjectCode = m.ProjectCode;
                        modelList.Add(mtmp);
                    }

                }// end for
                excelWBook.Close();

                //// split models by the status
                //List<Model> listCompletedPjts = modelList.FindAll(delegate(Model mm) { return (mm.ModelCA.ProjectStatus.Equals("COMPLETED")); });
                //List<Model> listDroppedPjts = modelList.FindAll(delegate(Model mm) { return (mm.ModelCA.ProjectStatus.Equals("DROPPED")); });
                //List<Model> listHoldPjts = modelList.FindAll(delegate(Model mm) { return (mm.ModelCA.ProjectStatus.Equals("HOLD")); });
                //List<Model> listECOPjts = modelList.FindAll(delegate(Model mm) { return (mm.ModelCA.ProjectStatus.Equals("ECO")); });
                //List<Model> listWaitPjts = modelList.FindAll(delegate(Model mm) { return (mm.ModelCA.ProjectStatus.Equals("WAIT")); });
                //List<Model> listRunningPjts = modelList.FindAll(delegate(Model mm) { return (mm.ModelCA.ProjectStatus.Equals("RUNNING")); });

                /*********************************
                 * * Print Data to Excel Files * *
                 * *******************************/
                // open the spreadsheet to store the result of the calculus
                Excel.Workbook xlWorkBook = null;
                xlWorkBook = excelApp.Workbooks.Open(fileDestPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                Excel.Sheets xlWorkSheets = (Excel.Sheets)xlWorkBook.Worksheets;
                Excel.Worksheet xlWorkSheet = xlWorkSheets.get_Item(1);

                /******************************
                 * Man Month for All Projects *
                 ******************************/
                // write date to the spreadsheet
                xlWorkSheet.Name = "BASIC MAN MONTH REPORT";
                Excel.Range xlRange = xlWorkSheet.get_Range("A1", "M" + modelList.Count.ToString());
                Int32 row = 2;
                // set column heads
                xlRange = xlWorkSheet.get_Range("A1", "A1");
                xlRange.Value = "Model";
                xlRange = xlWorkSheet.get_Range("B1", "B1");
                xlRange.Value = "Country";
                xlRange = xlWorkSheet.get_Range("C1", "C1");
                xlRange.Value = "Carrier";
                xlRange = xlWorkSheet.get_Range("D1", "D1");
                xlRange.Value = "Status";
                xlRange = xlWorkSheet.get_Range("E1", "E1");
                xlRange.Value = "SW";
                xlRange = xlWorkSheet.get_Range("F1", "F1");
                xlRange.Value = "UFC";
                xlRange = xlWorkSheet.get_Range("G1", "G1");
                xlRange.Value = "LSI";
                xlRange = xlWorkSheet.get_Range("H1", "H1");
                xlRange.Value = "PVG R&D";
                xlRange = xlWorkSheet.get_Range("I1", "I1");
                xlRange.Value = "PVG NOT R&D";
                xlRange = xlWorkSheet.get_Range("J1", "J1");
                xlRange.Value = "TAM";
                xlRange = xlWorkSheet.get_Range("K1", "K1");
                xlRange.Value = "BRISA";
                xlRange = xlWorkSheet.get_Range("L1", "L1");
                xlRange.Value = "Subsidiary";
                xlRange = xlWorkSheet.get_Range("M1", "M1");
                xlRange.Value = "Code";

                // fulfill all the rows
                foreach (Model model in modelList)
                {
                    // flags if the project is a R&D Project or not according MCT rules
                    bool rdProject = false;

                    xlRange = xlWorkSheet.get_Range("A" + row.ToString(), "A" + row.ToString());
                    xlRange.Value = model.ModelCode.ToString();
                    xlRange = xlWorkSheet.get_Range("B" + row.ToString(), "B" + row.ToString());
                    xlRange.Value = model.ModelCA.Country.ToString();
                    xlRange = xlWorkSheet.get_Range("C" + row.ToString(), "C" + row.ToString());
                    xlRange.Value = model.ModelCA.CarrierName.ToString();
                    xlRange = xlWorkSheet.get_Range("D" + row.ToString(), "D" + row.ToString());
                    xlRange.Value = model.ModelCA.ProjectStatus;

                    if (model.ModelCA.ListReportedHours.Find(delegate(ReportHours rh) { return (rh.TeamName.Equals("SW") || rh.TeamName.Equals("UFC") || rh.TeamName.Equals("LSI")); })!=null) 
                        rdProject = true;

                    foreach (ReportHours rh in model.ModelCA.ListReportedHours)
                    {
                        if(rh.TeamName.Equals("SW"))
                            xlRange = xlWorkSheet.get_Range("E" + row.ToString(), "E" + row.ToString());
                        else if (rh.TeamName.Equals("UFC"))
                            xlRange = xlWorkSheet.get_Range("F" + row.ToString(), "F" + row.ToString());
                        else if (rh.TeamName.Equals("LSI"))
                            xlRange = xlWorkSheet.get_Range("G" + row.ToString(), "G" + row.ToString());
                        else if (rh.TeamName.Equals("PVG"))
                        {
                            if(rdProject)
                                xlRange = xlWorkSheet.get_Range("H" + row.ToString(), "H" + row.ToString());
                            else
                                xlRange = xlWorkSheet.get_Range("I" + row.ToString(), "I" + row.ToString());
                        }
                        else if (rh.TeamName.Equals("TAM"))
                            xlRange = xlWorkSheet.get_Range("J" + row.ToString(), "J" + row.ToString());
                        else if (rh.TeamName.Equals("BRISA"))
                            xlRange = xlWorkSheet.get_Range("K" + row.ToString(), "K" + row.ToString());
                        
                        xlRange.NumberFormat = "0.000";
                        xlRange.Value = rh.ReportedTime;
                    }

                    xlRange = xlWorkSheet.get_Range("L" + row.ToString(), "L" + row.ToString());
                    xlRange.Value = model.ModelCA.Subsidiary;
                    xlRange = xlWorkSheet.get_Range("M" + row.ToString(), "M" + row.ToString());
                    xlRange.Value = model.ProjectCode;

                    row++;
                }
                xlWorkBook.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message + "\n\nLinha: " + linha + "\nPlanilha: " + planilha);
            }
            finally
            {
                // close Excel Application
                msg += "\n" + @"FIM: " + DateTime.Now.ToString("hh:mm:ss");
                MessageBox.Show(msg);
                excelApp.Quit();
            }
        }

        public void makeHRDelayCalculus(String filePath, String fileDestPath, String fileDiffWorkingTimeSheet, Object[] selectedWorkSheets, Object[] selectedColumnIndexes)
        {
            Excel._Workbook excelWBook;
            Excel._Worksheet excelWSheet;
            Excel.Range cellsRange;
            Object obj = null;

            int linhaLida = 0;
            int planilhaLida = 0;
            List<Employee> atrasadasMais1Hora = new List<Employee>();
            List<Employee> atrasadasMenos1Hora = new List<Employee>();
            List<Employee> emTempo = new List<Employee>();
            List<Employee> peopleDifferWorkingTime = new List<Employee>();
            DateTime arriveTime = new DateTime();

            List<Employee> atrasados15Minutos = new List<Employee>();
            List<Employee> atrasados20Minutos = new List<Employee>();
            List<Employee> atrasados30Minutos = new List<Employee>();
            List<Employee> atrasados40Minutos = new List<Employee>();
            List<Employee> atrasados50Minutos = new List<Employee>();
            List<Employee> atrasados60Minutos = new List<Employee>();

            // get the emnumerator of elements in each array
            System.Collections.IEnumerator numberSelectedWorkSheetProcess = selectedWorkSheets.GetEnumerator();

            try
            {
                // open the spreadsheet with the source data
                excelWBook = excelApp.Workbooks.Open(filePath, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, false, Type.Missing, Type.Missing, false, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                // move to the first spreadsheet in the document
                numberSelectedWorkSheetProcess.Reset();

                // read all sreapsheets existent in the Excel File
                while (numberSelectedWorkSheetProcess.MoveNext())
                {
                    // get the content of the current SpreadSheet
                    excelWSheet = (Excel.Worksheet)excelWBook.Worksheets.get_Item((Int32)numberSelectedWorkSheetProcess.Current + 1);
                    planilhaLida = (Int32)numberSelectedWorkSheetProcess.Current + 1;

                    // get the range of used columns and rows
                    cellsRange = excelWSheet.UsedRange;
                    String cell = null;
                    String name = null;

                    Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");

                    // collect the data
                    for (int rCount = 2; rCount < cellsRange.Rows.Count; rCount++)
                    {
                        linhaLida = rCount;
                        String numberFormat = null;
                        Employee p = new Employee();

                        cell = ("C" + rCount);
                        numberFormat = cellsRange.get_Range(cell, cell).NumberFormat;
                        if (numberFormat.Contains(@"yy"))
                        {
                            obj = cellsRange.get_Range(cell, cell).Value2;

                            DateTime date = DateTime.FromOADate(Convert.ToDouble(obj));
                            if (date.DayOfWeek.ToString().ToUpper().Equals("SUNDAY") || date.DayOfWeek.ToString().ToUpper().Equals("SATURDAY"))
                                continue;

                            p.Date = date;
                        }
                        else
                            continue;

                        cell = ("D" + rCount);
                        numberFormat = cellsRange.get_Range(cell, cell).NumberFormat;
                        if (numberFormat.Contains(@"h:mm"))
                        {
                            obj = cellsRange.get_Range(cell, cell).Value2;
                            p.Arrival = DateTime.FromOADate(Convert.ToDouble(obj));
                            p.Arrival = new DateTime(p.Date.Year, p.Date.Month, p.Date.Day, p.Arrival.Hour, p.Arrival.Minute, 0);
                        }

                        cell = ("G" + rCount);
                        numberFormat = cellsRange.get_Range(cell, cell).NumberFormat;
                        if (numberFormat.Contains(@"h:mm"))
                        {
                            obj = cellsRange.get_Range(cell, cell).Value2;
                            p.Left = DateTime.FromOADate(Convert.ToDouble(obj));
                            p.Left = new DateTime(p.Date.Year, p.Date.Month, p.Date.Day, p.Left.Hour, p.Left.Minute, 0);
                        }

                        cell = ("B" + rCount);
                        name = cellsRange.get_Range(cell, cell).Value2;
                        if (name == null)
                            continue;

                        p.Name = name;

                        // fix date to be the same as the inputted by system
                        arriveTime = new DateTime(p.Date.Year, p.Date.Month, p.Date.Day, 8, 0, 0);

                        if (arriveTime.AddMinutes(60.0) < p.Arrival)
                            atrasadasMais1Hora.Add(p);
                        else if (p.Arrival > arriveTime.AddMinutes(15.0) && arriveTime.AddMinutes(60.0) >= p.Arrival)
                        {
                            atrasadasMenos1Hora.Add(p);
                            //if (p.Arrival > arriveTime.AddMinutes(10.0) && arriveTime.AddMinutes(15.0) >= p.Arrival)
                            //    atrasados15Minutos.Add(p);
                            if (p.Arrival > arriveTime.AddMinutes(15.0) && arriveTime.AddMinutes(20.0) >= p.Arrival)
                                atrasados20Minutos.Add(p);
                            else if (p.Arrival > arriveTime.AddMinutes(20.0) && arriveTime.AddMinutes(30.0) >= p.Arrival)
                                atrasados30Minutos.Add(p);
                            else if (p.Arrival > arriveTime.AddMinutes(30.0) && arriveTime.AddMinutes(40.0) >= p.Arrival)
                                atrasados40Minutos.Add(p);
                            else if (p.Arrival > arriveTime.AddMinutes(40.0) && arriveTime.AddMinutes(50.0) >= p.Arrival)
                                atrasados50Minutos.Add(p);
                            else if (p.Arrival > arriveTime.AddMinutes(50.0) && arriveTime.AddMinutes(60.0) >= p.Arrival)
                                atrasados60Minutos.Add(p);

                        }
                        else
                            emTempo.Add(p);
                    }
                }
                excelWBook.Close(true, null, null);

                // collect data from diferentiated working time
                excelWBook = excelApp.Workbooks.Open(fileDiffWorkingTimeSheet, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, false, Type.Missing, Type.Missing, false, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                excelWSheet = (Excel.Worksheet)excelWBook.Worksheets.get_Item(1);
                cellsRange = excelWSheet.UsedRange;

                // read all differentiated working time
                for (int rCount = 2; rCount < cellsRange.Rows.Count; rCount++)
                {
                    linhaLida = rCount;
                    Employee p = new Employee();
                    String cell = null;

                    cell = ("A" + rCount);
                    p.Name = (string)cellsRange.get_Range(cell, cell).Value2;
                    p.Name = p.Name.ToUpper();

                    cell = ("B" + rCount);
                    if (cellsRange.get_Range(cell, cell).NumberFormat.Contains(@"h:mm"))
                    {
                        obj = cellsRange.get_Range(cell, cell).Value2;
                        p.Arrival = DateTime.FromOADate(Convert.ToDouble(obj));
                    }

                    cell = ("C" + rCount);
                    if (cellsRange.get_Range(cell, cell).NumberFormat.Contains(@"h:mm"))
                    {
                        obj = cellsRange.get_Range(cell, cell).Value2;
                        p.Left = DateTime.FromOADate(Convert.ToDouble(obj));
                    }

                    peopleDifferWorkingTime.Add(p);

                }
                excelWBook.Close(true, null, null);

                foreach (Employee p in peopleDifferWorkingTime)
                {
                    DateTime timeLimit60Minutes = p.Arrival.AddMinutes(60);
                    DateTime timeLimit10Minutes = p.Arrival.AddMinutes(10);

                    List<Employee> pList = emTempo.FindAll(delegate(Employee pp) { return (pp.Name.Equals(p.Name));});

                    foreach (Employee pp in pList)
                    {
                        timeLimit60Minutes = new DateTime(pp.Date.Year, pp.Date.Month, pp.Date.Day, timeLimit60Minutes.Hour, timeLimit60Minutes.Minute, 0);
                        timeLimit10Minutes = new DateTime(pp.Date.Year, pp.Date.Month, pp.Date.Day, timeLimit10Minutes.Hour, timeLimit10Minutes.Minute, 0);

                        if (pp.Arrival > timeLimit60Minutes)
                        {
                            emTempo.Remove(pp);
                            atrasadasMais1Hora.Add(pp);
                        }
                        else if (pp.Arrival > timeLimit10Minutes && pp.Arrival <= timeLimit60Minutes)
                        {
                            emTempo.Remove(pp);
                            atrasadasMenos1Hora.Add(pp);
                        }
                    }

                }

                int totalGeral = (emTempo.Count + atrasadasMais1Hora.Count + atrasadasMenos1Hora.Count);
                int totalMais1Hora = 0;
                int totalMenos1Hora = 0;

                //List<Employee> pPaulo1H = atrasadasMais1Hora.FindAll(delegate(Employee pp) { return (pp.Name.Equals("PAULO RICARDO BATISTA MESQUITA")); });
                //List<Employee> pPaulo10M = atrasadasMenos1Hora.FindAll(delegate(Employee pp) { return (pp.Name.Equals("PAULO RICARDO BATISTA MESQUITA")); });

                String message = "Total counts: " + totalGeral;
                message += "\n\nDelay more than 60 minutes: " + atrasadasMais1Hora.Count;
                message += "\nPercentage: " + ((double)atrasadasMais1Hora.Count / totalGeral).ToString("0.00%");
                message += "\n\nDelay bewteen 11 and 60 minutes: " + atrasadasMenos1Hora.Count;
                message += "\nPercentage: " + ((double)atrasadasMenos1Hora.Count / totalGeral).ToString("0.00%"); ;

                message += "\n\nDelays up to 60 minutes:";
                //message += "\n11-15: " + atrasados15Minutos.Count;
                message += "\n16-20: " + atrasados20Minutos.Count;
                message += "\n21-30: " + atrasados30Minutos.Count;
                message += "\n31-40: " + atrasados40Minutos.Count;
                message += "\n41-50: " + atrasados50Minutos.Count;
                message += "\n51-60: " + atrasados60Minutos.Count;

                MessageBox.Show(message);

                // Split by Month
                for (int month = 1; month <= 12; month++)
                {
                    // count marked arrival time
                    totalGeral = emTempo.FindAll(delegate(Employee pp) { return (pp.Date.Month == month); }).Count;

                    // count delayed > 60 minutes
                    totalMais1Hora = atrasadasMais1Hora.FindAll(delegate(Employee pp) { return (pp.Date.Month == month); }).Count;

                    // count delayed from 11 up to 60 minutes
                    totalMenos1Hora = atrasadasMenos1Hora.FindAll(delegate(Employee pp) { return (pp.Date.Month == month); }).Count;

                    // count delayed in time intervals
                    //int total1115 = atrasados15Minutos.FindAll(delegate(Employee pp) { return (pp.Date.Month == month); }).Count;
                    int total1620 = atrasados20Minutos.FindAll(delegate(Employee pp) { return (pp.Date.Month == month); }).Count;
                    int total2130 = atrasados30Minutos.FindAll(delegate(Employee pp) { return (pp.Date.Month == month); }).Count;
                    int total3140 = atrasados40Minutos.FindAll(delegate(Employee pp) { return (pp.Date.Month == month); }).Count;
                    int total4150 = atrasados50Minutos.FindAll(delegate(Employee pp) { return (pp.Date.Month == month); }).Count;
                    int total5160 = atrasados60Minutos.FindAll(delegate(Employee pp) { return (pp.Date.Month == month); }).Count;

                    totalGeral += totalMais1Hora + totalMenos1Hora;

                    message = "Total counts: " + totalGeral + " for Month " + month;
                    message += "\n\nDelay more than 60 minutes: " + totalMais1Hora;
                    message += "\nPercentage: " + ((double)totalMais1Hora / totalGeral).ToString("0.00%");
                    message += "\n\nDelay bewteen 11 and 60 minutes: " + totalMenos1Hora;
                    message += "\nPercentage: " + ((double)totalMenos1Hora / totalGeral).ToString("0.00%"); ;

                    message += "\n\nDelays up to 60 minutes:";
                    //message += "\n11-15: " + total1115 + " (" + ((double)total1115 / totalMenos1Hora).ToString("0.00%") +")";
                    message += "\n16-20: " + total1620 + " (" + ((double)total1620 / totalMenos1Hora).ToString("0.00%") + ")";
                    message += "\n21-30: " + total2130 + " (" + ((double)total2130 / totalMenos1Hora).ToString("0.00%") + ")";
                    message += "\n31-40: " + total3140 + " (" + ((double)total3140 / totalMenos1Hora).ToString("0.00%") + ")";
                    message += "\n41-50: " + total4150 + " (" + ((double)total4150 / totalMenos1Hora).ToString("0.00%") + ")";
                    message += "\n51-60: " + total5160 + " (" + ((double)total5160 / totalMenos1Hora).ToString("0.00%") + ")";

                    //MessageBox.Show(message);

                }

                // open the spreadsheet to store the result of the calculus
                Excel.Workbook xlWorkBook = null;
                xlWorkBook = excelApp.Workbooks.Open(fileDestPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                Excel.Sheets xlWorkSheets = (Excel.Sheets)xlWorkBook.Worksheets;
                Excel.Worksheet xlWorkSheet = xlWorkSheets.get_Item(1);

                // write date to the spreadsheet
                Excel.Range xlRange = xlWorkSheet.get_Range("A1", "L" + 100);

                // define head columns
                xlRange = xlWorkSheet.get_Range("A1", "A1");
                xlRange.Value = "Year";
                xlRange = xlWorkSheet.get_Range("B1", "B1");
                xlRange.Value = "2011";

                xlRange = xlWorkSheet.get_Range("A3", "A3");
                xlRange.Value = "Overal Results For Delays";

                xlRange = xlWorkSheet.get_Range("A4", "A4");
                xlRange.Value = "Total counts";
                xlRange = xlWorkSheet.get_Range("B4", "B4");
                xlRange.Value = emTempo.Count + atrasadasMais1Hora.Count + atrasadasMenos1Hora.Count;

                xlRange = xlWorkSheet.get_Range("A5", "A5");
                xlRange.Value = "More than 60 minutes";
                xlRange = xlWorkSheet.get_Range("B5", "B5");
                xlRange.Value = atrasadasMais1Hora.Count;
                xlRange = xlWorkSheet.get_Range("C5", "C5");
                xlRange.NumberFormat = "0.00%";
                xlRange.Value = ((double)atrasadasMais1Hora.Count / (atrasadasMais1Hora.Count+atrasadasMenos1Hora.Count+emTempo.Count));

                xlRange = xlWorkSheet.get_Range("E5", "E5");
                xlRange.Value = "Less than 60 minutes";
                xlRange = xlWorkSheet.get_Range("F5", "F5");
                xlRange.Value = atrasadasMenos1Hora.Count;
                xlRange = xlWorkSheet.get_Range("G5", "G5");
                xlRange.NumberFormat = "0.00%";
                xlRange.Value = ((double)atrasadasMenos1Hora.Count / (atrasadasMais1Hora.Count + atrasadasMenos1Hora.Count + emTempo.Count));

                xlRange = xlWorkSheet.get_Range("A7", "A7");
                xlRange.Value = "Results By Month";


                xlRange = xlWorkSheet.get_Range("A8", "A8");
                xlRange.Value = "Month";
                xlRange = xlWorkSheet.get_Range("B8", "B8");
                xlRange.Value = "Counts";
                xlRange = xlWorkSheet.get_Range("C8", "C8");
                xlRange.Value = "Delays > 1 hour";
                xlRange = xlWorkSheet.get_Range("D8", "D8");
                xlRange.Value = "%";
                xlRange = xlWorkSheet.get_Range("E8", "E8");
                xlRange.Value = "Delays < 1 hour";
                xlRange = xlWorkSheet.get_Range("F8", "F8");
                xlRange.Value = "%";
                //xlRange = xlWorkSheet.get_Range("G8", "G8");
                //xlRange.Value = "11-15";
                xlRange = xlWorkSheet.get_Range("H8", "H8");
                xlRange.Value = "16-20";
                xlRange = xlWorkSheet.get_Range("I8", "I8");
                xlRange.Value = "21-30";
                xlRange = xlWorkSheet.get_Range("J8", "J8");
                xlRange.Value = "31-40";
                xlRange = xlWorkSheet.get_Range("K8", "K8");
                xlRange.Value = "41-50";
                xlRange = xlWorkSheet.get_Range("L8", "L8");
                xlRange.Value = "51-60";

                Int32 row = 9;
                String[] months = { "", "JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC" };
                for (int month = 1; month <= 12; month++)
                {
                    totalGeral = emTempo.FindAll(delegate(Employee pp) { return (pp.Date.Month == month); }).Count;
                    totalGeral += atrasadasMais1Hora.FindAll(delegate(Employee pp) { return (pp.Date.Month == month); }).Count;
                    totalGeral += atrasadasMenos1Hora.FindAll(delegate(Employee pp) { return (pp.Date.Month == month); }).Count;

                    totalMais1Hora = atrasadasMais1Hora.FindAll(delegate(Employee pp) { return (pp.Date.Month == month); }).Count;
                    totalMenos1Hora = atrasadasMenos1Hora.FindAll(delegate(Employee pp) { return (pp.Date.Month == month); }).Count;

                    xlRange = xlWorkSheet.get_Range("A" + row.ToString(), "A" + row.ToString());
                    xlRange.Value = months[month];
                    xlRange = xlWorkSheet.get_Range("B" + row.ToString(), "B" + row.ToString());
                    xlRange.Value = totalGeral;
                    xlRange = xlWorkSheet.get_Range("C" + row.ToString(), "C" + row.ToString());
                    xlRange.Value = totalMais1Hora;
                    xlRange = xlWorkSheet.get_Range("D" + row.ToString(), "D" + row.ToString());
                    xlRange.NumberFormat = "0.00%";
                    xlRange.Value = ((double)totalMais1Hora / totalGeral);
                    xlRange = xlWorkSheet.get_Range("E" + row.ToString(), "E" + row.ToString());
                    xlRange.Value = totalMenos1Hora;
                    xlRange = xlWorkSheet.get_Range("F" + row.ToString(), "F" + row.ToString());
                    xlRange.NumberFormat = "0.00%";
                    xlRange.Value = ((double)totalMenos1Hora / totalGeral);
                    //xlRange = xlWorkSheet.get_Range("G" + row.ToString(), "G" + row.ToString());
                    //xlRange.Value = atrasados15Minutos.FindAll(delegate(Employee pp) { return (pp.Date.Month == month); }).Count;
                    xlRange = xlWorkSheet.get_Range("H" + row.ToString(), "H" + row.ToString());
                    xlRange.Value = atrasados20Minutos.FindAll(delegate(Employee pp) { return (pp.Date.Month == month); }).Count;
                    xlRange = xlWorkSheet.get_Range("I" + row.ToString(), "I" + row.ToString());
                    xlRange.Value = atrasados30Minutos.FindAll(delegate(Employee pp) { return (pp.Date.Month == month); }).Count;
                    xlRange = xlWorkSheet.get_Range("J" + row.ToString(), "J" + row.ToString());
                    xlRange.Value = atrasados40Minutos.FindAll(delegate(Employee pp) { return (pp.Date.Month == month); }).Count;
                    xlRange = xlWorkSheet.get_Range("K" + row.ToString(), "K" + row.ToString());
                    xlRange.Value = atrasados50Minutos.FindAll(delegate(Employee pp) { return (pp.Date.Month == month); }).Count;
                    xlRange = xlWorkSheet.get_Range("L" + row.ToString(), "L" + row.ToString());
                    xlRange.Value = atrasados60Minutos.FindAll(delegate(Employee pp) { return (pp.Date.Month == month); }).Count;

                    row++;
                }

                xlWorkBook.Close();
                MessageBox.Show("Process Finished");


                //MessageBox.Show("Continua");

                //// open the spreadsheet to store the result of the calculus
                //xlWorkBook = null;
                //xlWorkBook = excelApp.Workbooks.Open(fileDestPath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                //xlWorkSheets = (Excel.Sheets)xlWorkBook.Worksheets;
                //xlWorkSheet = xlWorkSheets.get_Item(1);
                //xlRange = xlWorkSheet.get_Range("A1", "E" + atrasadasMenos1Hora.Count);

                //xlRange = xlWorkSheet.get_Range("A1", "A1");
                //xlRange.Value = "Name";
                //xlRange = xlWorkSheet.get_Range("B1", "B1");
                //xlRange.Value = "Date";
                //xlRange = xlWorkSheet.get_Range("C1", "C1");
                //xlRange.Value = "Arrival Time";
                ////xlRange = xlWorkSheet.get_Range("D1", "D1");
                ////xlRange.Value = "Left Time";
                ////xlRange = xlWorkSheet.get_Range("E1", "E1");
                ////xlRange.Value = "Hours at Office";

                //row = 2;

                //foreach (Employee p in atrasadasMenos1Hora)
                //{
                //    xlRange = xlWorkSheet.get_Range("A" + row.ToString(), "A" + row.ToString());
                //    xlRange.Value = p.Name;
                //    xlRange = xlWorkSheet.get_Range("B" + row.ToString(), "B" + row.ToString());
                //    xlRange.Value = p.Date.ToString("yyyy/MM/dd");
                //    xlRange = xlWorkSheet.get_Range("C" + row.ToString(), "C" + row.ToString());
                //    xlRange.Value = p.Arrival.ToString("hh:mm");
                //    //xlRange = xlWorkSheet.get_Range("D" + row.ToString(), "D" + row.ToString());
                //    //xlRange.Value = p.Left.ToString("hh:mm");
                //    //xlRange = xlWorkSheet.get_Range("E" + row.ToString(), "E" + row.ToString());
                //    //xlRange.Value = p.WorkingHours.ToString("hh:mm");

                //    row++;
                //}

                //xlWorkBook.Close();

            }
            catch (Exception e)
            {
                String msg = "Work Sheet: " + planilhaLida;
                msg += "\nLine: " + linhaLida;
                msg += "\n" + e.Message;
                MessageBox.Show(msg);
            }
            finally
            {
                // close Excel Application
                excelApp.Quit();
            }
        }

        public void readHeadColumnsOfWorkSheet(String filePath, Object[] selectedWorkSheets)
        {
            Excel._Workbook excelWBook;
            Excel._Worksheet excelWSheet;
            columnHeads = new List<String>();
            // get the number of WorkSheets to process
            System.Collections.IEnumerator numberWorkSheetsProcess = selectedWorkSheets.GetEnumerator();

            try
            {
                // open Excel file to process
                excelWBook = excelApp.Workbooks.Open(filePath, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, false, Type.Missing, Type.Missing, false, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                // read all sreapsheets existent in the Excel File
                while (numberWorkSheetsProcess.MoveNext())
                {
                    //excelWSheet = (Excel.Worksheet)excelWBook.Worksheets.get_Item(cWSheet);
                    Int32 numSP = (Int32)numberWorkSheetsProcess.Current + 1;
                    excelWSheet = (Excel.Worksheet)excelWBook.Worksheets.get_Item(numSP);
                    Excel.Range cellsRange = excelWSheet.UsedRange;

                    // collect all column names in the spreadsheet
                    for (int cCount = 1; cCount <= cellsRange.Columns.Count; cCount++)
                        columnHeads.Add((String)(cellsRange.Cells[1, cCount] as Excel.Range).Value2);

                    break; // in this moment, just one spreadsheet is needed

                }
                excelWBook.Close(true, null, null);

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            finally
            {
                // close Excel Application
                excelApp.Quit();
            }
        }

        public List<CountryCode> getCountryCodes(String filePath)
        {
            Excel._Workbook excelWBook;
            Excel._Worksheet excelWSheet;
            List<CountryCode> listCountryCodes = new List<CountryCode>();

            try
            {
                // open Excel file to process
                excelWBook = excelApp.Workbooks.Open(filePath, Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, false, Type.Missing, Type.Missing, false, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                // get the content of the current SpreadSheet
                excelWSheet = (Excel.Worksheet)excelWBook.Worksheets.get_Item(1);
                // get the used ranges in the excel spreadsheet
                Excel.Range excelRange = excelWSheet.UsedRange;
                // read all sreapsheets existent in the Excel File
                CountryCode cc = null;
                for(int row = 2; row <= excelRange.Rows.Count; row++)
                {
                    cc = new CountryCode();                    
                    cc.Code = (string)excelRange.get_Range("A" + row, "A" + row).Value2;

                    String[] tmp = cc.Code.Split('.');
                    cc.Code = tmp[1];
                    cc.Code = cc.Code.Substring(1, 3);

                    cc.Country = (string)excelRange.get_Range("B" + row, "B" + row).Value2;
                  
                    cc.Carrier = (string)excelRange.get_Range("C" + row, "C" + row).Value2;
                    if ((cc.Country == null || cc.Country.Equals("NO COUNTRY NAME") || cc.Country.Equals("VENEZUELA")) && (cc.Carrier == null || cc.Carrier.Equals("NO CARRIER NAME")))
                    {
                        cc.Country = "VENEZUELA";
                        cc.Carrier = "OPEN";
                    }
                    if (listCountryCodes.Count > 0)
                        if (listCountryCodes.Find(delegate(CountryCode c) { return (c.Code == cc.Code); }) != null)
                            continue;

                    listCountryCodes.Add(cc);

                }
                excelWBook.Close(true, null, null);

                return listCountryCodes;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return listCountryCodes;
            }
            finally
            {
                // close Excel Application
                excelApp.Quit();
            }
        }

        public List<ProjectDB> ProjectList
        {
            get
            {
                return this.projectList;
            }
        }

        public List<String> ColumnHeads
        {
            get
            {
                return columnHeads;
            }
        }

        public List<String> WorkSheetNames
        {
            get
            {
                return wSheets;
            }
        }

        public String[,] CellValues
        {
            get
            {
                return cellValues;
            }
        }
    }
}

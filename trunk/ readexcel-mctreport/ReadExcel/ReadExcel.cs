using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace ReadExcel
{
    public class ReadExcel
    {
        private Excel.Application excelApp;
        private Excel._Workbook excelWB;
        private Excel._Worksheet excelWS;

        public ReadExcel()
        {
            excelApp = new Excel.Application();
        }

        private void processInformedWorkBooks(String filePath)
        {
            try
            {
                // try to open workbook in the selected file
                excelWB = excelApp.Workbooks.Open(filePath, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, true, Type.Missing, Type.Missing, false, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                // try to read the worksheet with the data, supposing it is in the first worksheet
                excelWS = (Excel.Worksheet)excelWB.Worksheets.get_Item(1);
            }
            catch (Exception e)
            {
                // show error message
                String msg = " \n\nwhen processing File " + filePath + "\nand Worksheet" + excelWS.Name;
                MessageBox.Show(e.Message + msg, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                // close the workbook and leave Excel application
                excelWB.Close();
                excelApp.Quit();
            }
        }

        //public 
    }
}

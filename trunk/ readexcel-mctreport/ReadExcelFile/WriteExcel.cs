using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace ReadExcelFile
{
    class WriteExcel
    {
        Excel.Application excelApp;

        public WriteExcel()
        {
            excelApp = new Excel.Application();
        }

        //public void createExcelFile(String filePath, List<Project> projectList)
        //{
        //    Excel._Workbook excelWBook;
        //    Excel.Workbooks excelWBooks;
        //    Excel._Worksheet excelWSheet;
        //    Excel.Sheets excelWSheets;
        //    Excel.Range excelRange;
        //    Int32 rCount = 2;

        //    try
        //    {
        //        // Start the workbook in Excel
        //        excelWBooks = excelApp.Workbooks;
        //        excelWBook = (Excel._Workbook)(excelWBooks.Add(Type.Missing));

        //        // Add data to head table cells in the first spreadsheeet in the new file
        //        excelWSheets = (Excel.Sheets)excelWBook.Worksheets;
        //        excelWSheet = (Excel._Worksheet)(excelWSheets.get_Item(1));
        //        excelRange = excelWSheet.get_Range("A1", Type.Missing);
        //        excelRange.Value = "Model";
        //        excelRange = excelWSheet.get_Range("B1", Type.Missing);
        //        excelRange.Value = "Carrier";
        //        excelRange = excelWSheet.get_Range("C1", Type.Missing);
        //        excelRange.Value = "Country";
        //        excelRange = excelWSheet.get_Range("D1", Type.Missing);
        //        excelRange.Value = "Quantity";

        //        // fulfill all cells with project data
        //        foreach (Project project in projectList)
        //        {
        //            // write project name and number of CAs
        //            excelRange = excelWSheet.get_Range("A" + rCount.ToString(), Type.Missing);
        //            excelRange.Value = project.Name;
        //            excelRange = excelWSheet.get_Range("D" + rCount.ToString(), Type.Missing);
        //            excelRange.Value = project.Quantity;

        //            foreach (CA ca in project.CAProjects)
        //            {
        //                // write carrier name in the spreadsheet
        //                excelRange = excelWSheet.get_Range("B" + rCount.ToString(), Type.Missing);
        //                excelRange.Value = ca.CACarrier;

        //                // write country name in the spreadsheet
        //                excelRange = excelWSheet.get_Range("C" + rCount.ToString(), Type.Missing);
        //                excelRange.Value = ca.CACountry;
                        
        //                // update rCount
        //                rCount++;
        //            }

        //        }

        //        // save data to excel spreadsheet
        //        excelWBook.SaveAs(filePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
        //    }
        //    catch (Exception e)
        //    {
        //        MessageBox.Show(e.Message);
        //    }
        //    finally
        //    {
        //        // quit Excel
        //        excelApp.Quit();
        //    }

        //}

        public void createExcelFile(String filePath, Object[] headColumns, String[,] cellValues)
        {
            Excel._Workbook excelWBook;
            Excel.Workbooks excelWBooks;
            Excel._Worksheet excelWSheet;
            Excel.Sheets excelWSheets;
            Excel.Range excelRange;
            Int32 numCols = 1;
            Int32 initialCharCode = 65;

            // get the list of column heads
            System.Collections.IEnumerator columnHeads = headColumns.GetEnumerator();

            String[] columnsChars = new String[26];
            for (int i = 0; i < 26; i++)
                columnsChars[i] = Convert.ToChar(initialCharCode + i).ToString();

            try
            {
                // Start the workbook in Excel
                excelWBooks = excelApp.Workbooks;
                excelWBook = (Excel._Workbook)(excelWBooks.Add(Type.Missing));

                // Add data to head table cells in the first spreadsheeet in the new file
                excelWSheets = (Excel.Sheets)excelWBook.Worksheets;
                excelWSheet = (Excel._Worksheet)(excelWSheets.get_Item(1));
                excelRange = excelWSheet.UsedRange;

                // write column heads
                while (columnHeads.MoveNext())
                {
                    excelRange = excelWSheet.get_Range(columnsChars[numCols - 1].ToString() + "1", Type.Missing);
                    excelRange.Value = columnHeads.Current.ToString();
                    numCols++;
                }

                // process all lines in the bi-dimensional array to write to Excel file
                for (int row = 2; row <= cellValues.GetLength(0); row++)
                    // process all columns in the secone dimension of the array
                    for (int column = 1; column < numCols; column++)
                    {
                        // write data to the cell
                        excelRange = excelWSheet.get_Range(columnsChars[column - 1].ToString() + row.ToString(), Type.Missing);
                        excelRange.Value = cellValues[row - 1, column - 1];
                    }

                Object[] columns = new Object[headColumns.Length];
                for(int i = 0; i < headColumns.Length; i++)
                    columns[i] = columnsChars[i].ToString() + "1";

                //excelRange.RemoveDuplicates(columns);
                // save data to excel spreadsheet
                excelWBook.SaveAs(filePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            finally
            {
                // quit Excel
                excelApp.Quit();
            }

        }


    }
}

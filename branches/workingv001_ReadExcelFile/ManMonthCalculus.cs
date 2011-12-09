using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace ReadExcelFile
{
    class ManMonthCalculus
    {
        Excel.Application excelApp = new Excel.Application();
        private Model model;
        private List<Model> listModels;

        public ManMonthCalculus()
        {
            model = new Model();
            listModels = new List<Model>();
        }

        public void calculateManMonth(String filePath, Object[] headColumns, String[,] cellValues)
        {
            Excel._Workbook excelWBook;
            Excel.Workbooks excelWBooks;
            Excel._Worksheet excelWSheet;
            Excel.Sheets excelWSheets;
            Excel.Range excelRange;
            Int32 numCols = cellValues.GetUpperBound(1);
            Int32 numRows = cellValues.GetUpperBound(0);
            Int32 initialCharCode = 65;
            String[] tmp = null;

            // get the list of column heads
            System.Collections.IEnumerator columnHeads = headColumns.GetEnumerator();

            String[] columnsChars = new String[26];
            for (int i = 0; i < 26; i++)
                columnsChars[i] = Convert.ToChar(initialCharCode + i).ToString();

            // avoid duplicated projects and duplicated information
            for (int i = 0; i < cellValues.GetUpperBound(0); i++)
            {

            }

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
                {
                    // process all columns in the secone dimension of the array
                    for (int column = 1; column < numCols; column++)
                    {
                        // write data to the cell
                        excelRange = excelWSheet.get_Range(columnsChars[column - 1].ToString() + row.ToString(), Type.Missing);
                        if (column == 1)
                        {
                            tmp = cellValues[row - 1, column - 1].Split('.');
                            excelRange.Value = tmp[0];
                        }
                        else
                            excelRange.Value = cellValues[row - 1, column - 1];
                    }
                }

                Object[] columns = new Object[headColumns.Length];
                for (int i = 0; i < headColumns.Length; i++)
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

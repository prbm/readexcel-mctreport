using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Windows.Forms;

namespace ReadExcelFile
{
    public partial class frmReadExcelMain : Form
    {
        List<String> columHeads;
        private List<CountryCode> listCountryCodes;
        private String diffWorkingTimeSheet = null;
        private String exceptionList = null;

        public frmReadExcelMain()
        {
            InitializeComponent();
        }

        // select the original file to be processed
        private void btBrowseFile_Click(object sender, EventArgs e)
        {
            String fileTmp = null;
            OpenFileDialog oFD = new OpenFileDialog();
            oFD.InitialDirectory = "C:\\";
            oFD.Filter = "Excel files 1997-2003 (*.xls)|*.xls|Excel files 2007-2011 (*.xlsx)|*.xlsx";
            oFD.FilterIndex = 1;

            // if there is a file already selected, store it in memory
            if (tbFileOpenPath.Text.Trim().Length > 0)
                fileTmp = tbFileOpenPath.Text.Trim();

            try
            {
                // check if a new file was chosen
                if (oFD.ShowDialog() == DialogResult.OK)
                    tbFileOpenPath.Text = oFD.FileName.Trim();
                else
                    return;

                // if the select file was changed
                if (!tbFileOpenPath.Text.Trim().Equals(fileTmp))
                {
                    //  clean the list box fields
                    lbColumnHeads.Items.Clear();
                    lbColumnHeads.Refresh();
                    lbWorkSheets.Items.Clear();
                    lbWorkSheets.Refresh();

                    Cursor.Current = Cursors.WaitCursor;

                    // get the WorkSheets from the selected file
                    ExcelRead er = new ExcelRead();
                    er.getWorkSheets(oFD.FileName);

                    // fulfill the WorkSheet list with the list got from Excel File
                    if (er.WorkSheetNames.Count > 0)
                    {
                        if (lbWorkSheets.Items.Count > 0)
                            lbWorkSheets.Items.Clear();

                        foreach (String str in er.WorkSheetNames)
                            lbWorkSheets.Items.Add(str);

                        if (lbWorkSheets.Items.Count > 0)
                            lbWorkSheets.Refresh();
                    }

                    stStripProcessing.Text = null;
                }

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

        // process selected file
        private void btProcess_Click(object sender, EventArgs e)
        {
            // if screen items are not ok, stop processing
            if (!checkIfScreenItemsAreOk(0))
                return;

            // there are worksheets to be processed?
            if (lbWorkSheets.SelectedItems.Count > 0)
            {
                // check with user if all spreadsheetd have the same heads
                if ((lbWorkSheets.SelectedItems.Count > 1) && MessageBox.Show("The Head Columns are the same for all SpreadSheets?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                {
                    lbWorkSheets.Refresh();
                    lbWorkSheets.Focus();
                    return;
                }
                else 
                {
                    // Collect the indexes of the select WorkSheets
                    Object[] selectWorksheetsIndexes = new Object[lbWorkSheets.SelectedItems.Count];
                    lbWorkSheets.SelectedIndices.CopyTo(selectWorksheetsIndexes, 0);

                    // get the column heads from the Excel File
                    ExcelRead er = new ExcelRead();
                    Cursor.Current = Cursors.WaitCursor; // starts hourglass cursor
                    er.readHeadColumnsOfWorkSheet(tbFileOpenPath.Text, selectWorksheetsIndexes);
                    columHeads = er.ColumnHeads;

                    // if columns where found in the worksheet, mount a list of it
                    if (columHeads.Count > 0)
                    {
                        if (lbColumnHeads.Items.Count > 0)
                        {
                            lbColumnHeads.Items.Clear();
                            lbColumnHeads.Refresh();
                        }

                        foreach (String str in columHeads)
                            if(str!=null)
                                lbColumnHeads.Items.Add(str);

                        // refresh the screen with the columns head list
                        lbColumnHeads.Refresh();
                    }
                } // end if (MessageBox.Show("The Head Columns are the same for all SpreadSheets?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            }// end if (lbWorkSheets.SelectedItems.Count > 0)

            Cursor.Current = Cursors.Default;
        }

        // select the file to store the result of processing
        private void btFileDestination_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.InitialDirectory = "C:\\";
            sfd.Filter = "Excel files 1997-2003 (*.xls)|*.xls|Excel files 2007-2011 (*.xlsx)|*.xlsx";
            sfd.FilterIndex = 1;

            try
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                    tbFileDestination.Text = sfd.FileName;
                else
                    return;
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
            finally
            {
                stStripProcessing.Text = null;
            }
        }

        private void pbSelectedColumns_Click(object sender, EventArgs e)
        {
            // are the items ok before processing?
            if (!checkIfScreenItemsAreOk(1))
                return;

            String[,] values;
            ExcelRead er = new ExcelRead();

            Cursor.Current = Cursors.WaitCursor;

            // Collect the indexes of the select WorkSheets
            Object[] selectWorksheetsIndexes = new Object[lbWorkSheets.SelectedItems.Count];
            lbWorkSheets.SelectedIndices.CopyTo(selectWorksheetsIndexes, 0);
            // Collect the indexes of the select columns
            Object[] selectColumnsIndexes = new Object[lbColumnHeads.SelectedItems.Count];
            lbColumnHeads.SelectedIndices.CopyTo(selectColumnsIndexes, 0);

            er.processWorkSheets(tbFileOpenPath.Text, selectWorksheetsIndexes, selectColumnsIndexes);
            values = new String[er.CellValues.GetLength(0), er.CellValues.GetLength(1)];
            values = er.CellValues;

            er = null;

            // Collect the indexes of the select columns
            Object[] selectColumnsItems = new Object[lbColumnHeads.SelectedItems.Count];
            lbColumnHeads.SelectedItems.CopyTo(selectColumnsItems, 0);

            if (values.Length > 0)
            {
                WriteExcel we = new WriteExcel();
                we.createExcelFile(tbFileDestination.Text, selectColumnsItems, values);
                we = null;

                MessageBox.Show("File created on " + tbFileDestination.Text);
            }
            
            Cursor.Current = Cursors.Default;
           
        }

        private void ThreadProc(Object data)
        {
            String[] tmp = data.ToString().Split('$');
            Form form = new frmPreviewExcel(tmp[0], tmp[1]);
            Application.Run(form);

        }

        private void pbWSheetPreview_Click(object sender, EventArgs e)
        {
            System.Threading.Thread thread = new System.Threading.Thread(new System.Threading.ParameterizedThreadStart(ThreadProc));
            thread.Start(tbFileOpenPath.Text + "$" + lbWorkSheets.SelectedItem.ToString().Trim());
        }

        // calculate the Man Month in the selected spreadsheets
        private void pbGenManMonthReport_Click(object sender, EventArgs e)
        {
            // check if the items in the screen are ok to be processed
            if (!checkIfScreenItemsAreOk(1))
                return;

            ExcelRead er = new ExcelRead();

            Cursor.Current = Cursors.WaitCursor;

            // Collect the indexes of the select WorkSheets
            Object[] selectWorksheetsIndexes = new Object[lbWorkSheets.SelectedItems.Count];
            lbWorkSheets.SelectedIndices.CopyTo(selectWorksheetsIndexes, 0);
            // Collect the indexes of the select columns
            Object[] selectColumnsIndexes = new Object[lbColumnHeads.SelectedItems.Count];
            lbColumnHeads.SelectedIndices.CopyTo(selectColumnsIndexes, 0);

            er.calculateManMonth(tbFileOpenPath.Text, tbFileDestination.Text, exceptionList, selectWorksheetsIndexes, selectColumnsIndexes, listCountryCodes);

            Cursor.Current = Cursors.Default;
        }

        // select the file with the list of emplyees with differentiated working time 
        private void pbBrowseExceptionSheet_Click(object sender, EventArgs e)
        {
            OpenFileDialog oFD = new OpenFileDialog();
            oFD.InitialDirectory = "C:\\";
            oFD.Filter = "Excel files 1997-2003 (*.xls)|*.xls|Excel files 2007-2011 (*.xlsx)|*.xlsx";
            oFD.FilterIndex = 1;

            try
            {
                if (oFD.ShowDialog() == DialogResult.OK)
                    diffWorkingTimeSheet = oFD.FileName;
                else
                    return;

            }
            catch(Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        // make the report of delays
        private void pbHRDelayCalculus_Click(object sender, EventArgs e)
        {
            if (diffWorkingTimeSheet == null || diffWorkingTimeSheet.Length == 0)
            {
                MessageBox.Show("No WorkSheet with differentiated working time was selected. Do it first!");
                pbBrowseExceptionSheet.Focus();
                return;
            }

            // the items in the screen to be processed are ok?
            if (!checkIfScreenItemsAreOk(1))
                return;

            ExcelRead er = new ExcelRead();

            Cursor.Current = Cursors.WaitCursor;

            // Collect the indexes of the select WorkSheets
            Object[] selectWorksheetsIndexes = new Object[lbWorkSheets.SelectedItems.Count];
            lbWorkSheets.SelectedIndices.CopyTo(selectWorksheetsIndexes, 0);
            // Collect the indexes of the select columns
            Object[] selectColumnsIndexes = new Object[lbColumnHeads.SelectedItems.Count];
            lbColumnHeads.SelectedIndices.CopyTo(selectColumnsIndexes, 0);

            er.makeHRDelayCalculus(tbFileOpenPath.Text, tbFileDestination.Text, diffWorkingTimeSheet, selectWorksheetsIndexes, selectColumnsIndexes);

            Cursor.Current = Cursors.Default;

        }

        // select the file with the status list of working projects
        private void pbProjectStatusList_Click(object sender, EventArgs e)
        {
            OpenFileDialog oFD = new OpenFileDialog();
            oFD.InitialDirectory = "C:\\";
            oFD.Filter = "Excel files 1997-2003 (*.xls)|*.xls|Excel files 2007-2011 (*.xlsx)|*.xlsx";
            oFD.FilterIndex = 1;

            try
            {
                if (oFD.ShowDialog() == DialogResult.OK)
                    exceptionList = oFD.FileName;
                else
                    return;
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        // select the country codes to be processed
        private void button1_Click(object sender, EventArgs e)
        {
            String fileTmp = null;
            OpenFileDialog oFD = new OpenFileDialog();
            oFD.InitialDirectory = "C:\\";
            oFD.Filter = "Excel files 1997-2003 (*.xls)|*.xls|Excel files 2007-2011 (*.xlsx)|*.xlsx";
            oFD.FilterIndex = 1;

            try
            {
                if (oFD.ShowDialog() == DialogResult.OK)
                {
                    fileTmp = oFD.FileName;


                    ExcelRead er = new ExcelRead();
                    listCountryCodes = er.getCountryCodes(fileTmp);

                    Cursor.Current = Cursors.WaitCursor;
                }

            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }

        }

        // check if the items in the screen are OK before processing data
        private bool checkIfScreenItemsAreOk(int checkingLevel)
        {

            switch (checkingLevel)
            {
                //check if there are selected columns to be processed
                case 1:
                    if (lbColumnHeads.SelectedItems.Count == 0)
                    {
                        MessageBox.Show("There are no selected columns");
                        lbColumnHeads.Refresh();
                        lbColumnHeads.Focus();
                        return false;
                    }
                    break;
                // check if there are selected worksheets to be processed
                case 0:
                    if (lbWorkSheets.SelectedItems.Count == 0)
                    {
                        MessageBox.Show("No WorkSheets were selected!");
                        lbWorkSheets.Refresh();
                        lbWorkSheets.Focus();
                        return false;
                    }
                    break;
            }// end switch

            return true;
        }// end method checkIfScreenItemsAreOk

    }

}

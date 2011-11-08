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
        List<Project> projects;
        List<String> columHeads;
        private List<CountryCode> listCountryCodes;
        private String diffWorkingTimeSheet = null;
        private String exceptionList = null;

        public frmReadExcelMain()
        {
            InitializeComponent();
        }

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
                if (oFD.ShowDialog() == DialogResult.OK)
                    tbFileOpenPath.Text = oFD.FileName;

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

        private void btProcess_Click(object sender, EventArgs e)
        {

            if (lbWorkSheets.SelectedItems.Count == 0)
            {
                MessageBox.Show("No WorkSheets were selected!");
                lbWorkSheets.Refresh();
                lbWorkSheets.Focus();
                return;
            }

            // in this moment, does not allow more than 1 worksheet at a time
            if (lbWorkSheets.SelectedItems.Count > 0)
            {
                //MessageBox.Show("Operation not allowed for more then 1 SpreadSheet!");
                if ((lbWorkSheets.SelectedItems.Count > 1) && MessageBox.Show("The Head Columns are the same for all SpreadSheets?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                {
                    lbWorkSheets.Refresh();
                    lbWorkSheets.Focus();
                    return;
                }
                else // if (MessageBox.Show("The Head Columns are the same for all SpreadSheets?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
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
                        {
                            if(str!=null)
                                lbColumnHeads.Items.Add(str);
                        }

                        lbColumnHeads.Refresh();
                    }
                } // fim if (MessageBox.Show("The Head Columns are the same for all SpreadSheets?", "Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            }

            Cursor.Current = Cursors.Default;
        }

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
            if (lbColumnHeads.SelectedItems.Count == 0)
            {
                MessageBox.Show("There are no selected columns");
                lbColumnHeads.Refresh();
                lbColumnHeads.Focus();
                return;
            }

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

        private void pbGenManMonthReport_Click(object sender, EventArgs e)
        {
            if (lbColumnHeads.SelectedItems.Count == 0)
            {
                MessageBox.Show("There are no selected columns");
                lbColumnHeads.Refresh();
                lbColumnHeads.Focus();
                return;
            }

            ExcelRead er = new ExcelRead();

            Cursor.Current = Cursors.WaitCursor;

            // Collect the indexes of the select WorkSheets
            Object[] selectWorksheetsIndexes = new Object[lbWorkSheets.SelectedItems.Count];
            lbWorkSheets.SelectedIndices.CopyTo(selectWorksheetsIndexes, 0);
            // Collect the indexes of the select columns
            Object[] selectColumnsIndexes = new Object[lbColumnHeads.SelectedItems.Count];
            lbColumnHeads.SelectedIndices.CopyTo(selectColumnsIndexes, 0);

            er.getManMonthPVGTAM(tbFileOpenPath.Text, tbFileDestination.Text, exceptionList, selectWorksheetsIndexes, selectColumnsIndexes, listCountryCodes);

            Cursor.Current = Cursors.Default;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            String fileTmp = null;
            OpenFileDialog oFD = new OpenFileDialog();
            oFD.InitialDirectory = "C:\\";
            oFD.Filter = "Excel files 1997-2003 (*.xls)|*.xls|Excel files 2007-2011 (*.xlsx)|*.xlsx";
            oFD.FilterIndex = 1;

            //// if there is a file already selected, store it in memory
            //if (diffWorkingTimeSheet.Trim().Length > 0)
            //    fileTmp = diffWorkingTimeSheet.Trim();
            try
            {
                if (oFD.ShowDialog() == DialogResult.OK)
                    diffWorkingTimeSheet = oFD.FileName;

            }
            catch(Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

        private void pbHRDelayCalculus_Click(object sender, EventArgs e)
        {
            if (diffWorkingTimeSheet == null || diffWorkingTimeSheet.Length == 0)
            {
                MessageBox.Show("No WorkSheet with differentiated working time was selected. Do it first!");
                button2.Focus();
                return;
            }

            if (lbColumnHeads.SelectedItems.Count == 0)
            {
                MessageBox.Show("There are no selected columns");
                lbColumnHeads.Refresh();
                lbColumnHeads.Focus();
                return;
            }

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
                else if (oFD.ShowDialog() == DialogResult.Cancel)
                    return;
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }

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

    }

}

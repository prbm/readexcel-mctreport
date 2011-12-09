using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data;
using System.Data.OleDb;
using Microsoft.Win32;

namespace GetCountryCodes
{
    /// <summary>
    /// Interaction logic for frmCountryCarriersPreview.xaml
    /// </summary>
    public partial class frmCountryCarriersPreview : Window
    {
        public frmCountryCarriersPreview()
        {
            InitializeComponent();
        }

        public void pbBrowseClicked(object sender, RoutedEventArgs e)
        {
            // Declare variables for connection strings usage
            OleDbConnection odbConn = null;
            OleDbDataAdapter odbDataAdapter = null;

            String fileName;
            // select file to get the data to previsualize
            OpenFileDialog dlg = new OpenFileDialog();
            //dlg.DefaultExt = "*.xlsx";
            dlg.Filter = "Excel 1997-2003 (*.xls)|*.xls|Excel 2007-2010 (*.xlsx)|*.xlsx";
            dlg.FilterIndex = 2;

            // get the name and path of selected file
            Nullable<bool> result = dlg.ShowDialog();
            if (result == true)
                fileName = dlg.FileName;
            else
                return;

            try
            {
                // provide string connection for data provider
                String strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=Excel 8.0;";

                odbConn = new OleDbConnection(strConn);
                odbConn.Open();

                OleDbCommand cmdSelect = new OleDbCommand("SELECT * FROM [Plan1$];", odbConn);

                odbDataAdapter = new OleDbDataAdapter();
                DataTable dTable = new DataTable();
                odbDataAdapter.SelectCommand = cmdSelect;

                DataSet objDataSet = new DataSet();
                odbDataAdapter.Fill(objDataSet);
                gridExcelPreview.BeginInit();
                gridExcelPreview.ItemsSource = objDataSet.Tables[0].DefaultView;
                gridExcelPreview.Items.Refresh();
                gridExcelPreview.EndInit();

                //gridExcelPreview.SelectionMode = DataGridSelectionMode.Single;
                gridExcelPreview.SelectionUnit = DataGridSelectionUnit.Cell;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                // close open connections
                odbConn.Close();
                odbDataAdapter = null;
            }
        }


        private void pbGetCountryCarriers_Click(object sender, RoutedEventArgs e)
        {
            // local auxiliary variables
            String[] selectedHeaders = new string[gridExcelPreview.SelectedCells.Count];
            List<String> headers = new List<String>();
            int countColumns = 0;

            // get column headers
            foreach (DataGridCellInfo cell in gridExcelPreview.SelectedCells)
            {
                // avoid duplicated headers counting
                if (headers.Count > 0)
                {
                    if (!headers.Contains(cell.Column.Header.ToString()))
                        headers.Add(cell.Column.Header.ToString());
                    else
                        continue;
                }
                else
                    headers.Add(cell.Column.Header.ToString());

                selectedHeaders[countColumns++] = cell.Column.Header.ToString();
            }

            if (countColumns < 3)
            {
                MessageBox.Show("Not enough columns were selected to get Model code, Carrier and Country names", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                gridExcelPreview.Items.Refresh();
                gridExcelPreview.Focus();
                return;
            }

        }
    }
}

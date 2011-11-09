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
                String strConn = @"Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + fileName + 
                                 @"; Jet OLEDB:Engine Type=5;Extended Properties=Excel 8.0;";
                                       //";Extended Properties=Excel 8.0;";

                OleDbConnection odbConn = new OleDbConnection(strConn);
                odbConn.Open();

                OleDbCommand cmdSelect = new OleDbCommand(@"SELECT * FROM [Plan1$];", odbConn);


                OleDbDataAdapter odbDataAdapter = new OleDbDataAdapter();
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

                odbConn.Close();
                odbDataAdapter = null;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
            }
        }


        private void pbGetCountryCarriers_Click(object sender, RoutedEventArgs e)
        {
            String[] selectedHeaders = new string[gridExcelPreview.SelectedCells.Count];
            int countColumns = 0;

            foreach (DataGridCellInfo cell in gridExcelPreview.SelectedCells)
                selectedHeaders[countColumns++] = cell.Column.Header.ToString();

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

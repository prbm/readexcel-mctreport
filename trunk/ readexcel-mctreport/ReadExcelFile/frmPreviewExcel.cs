using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.OleDb;

namespace ReadExcelFile
{
    public partial class frmPreviewExcel : Form
    {
        String filePath;
        String spreadSheet;

        public frmPreviewExcel()
        {
            InitializeComponent();
            filePath = null;
        }

        public frmPreviewExcel(String filePath, String spreadSheet)
        {
            InitializeComponent();
            this.filePath = filePath;
            this.spreadSheet = spreadSheet;
        }

        private void frmPreviewExcel_Load(object sender, EventArgs e)
        {
            try
            {
                String strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                                       "Data Source=" + filePath + "; Jet OLEDB:Engine Type=5;Extended Properties=Excel 8.0;";

                OleDbConnection oDbConn = new OleDbConnection(strConnection);
                oDbConn.Open();
                OleDbCommand cmdSelect = new OleDbCommand(@"SELECT * FROM [" + spreadSheet + "$]", oDbConn);
                OleDbDataAdapter oDbDataAdapter = new OleDbDataAdapter();
                oDbDataAdapter.SelectCommand = cmdSelect;
                DataTable dTable = new DataTable();
                oDbDataAdapter.Fill(dTable);
                oDbConn.Close();
                oDbDataAdapter = null;

                dataGridSpreadSheetOverview.DataSource = dTable;

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void pbOk_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}

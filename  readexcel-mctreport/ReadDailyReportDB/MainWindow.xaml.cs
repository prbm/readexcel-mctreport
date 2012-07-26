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

namespace ReadDailyReportDB
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void pbConnectDB_Click(object sender, RoutedEventArgs e)
        {
            String msg = null;
            DailyReportDB drDB = new DailyReportDB();

            msg = "select sum(time_to_sec(effortxproject.time)) from effortxproject inner join (effort, daily) " +
                  "on (daily.idDaily = effort.idDaily and effortxproject.idEffort = effort.idEffort) " +
                  "where " +
                  "(daily.date>='2011-11-01' and daily.date<='2011-11-30')";

            if (drDB.openConnection() == true)
            {
                MessageBox.Show("Conectou", "Info", MessageBoxButton.OK, MessageBoxImage.Information);

                drDB.executeSelect(msg, "sum(time_to_sec(effortxproject.time))");

                if(drDB.closeConnection()==true)
                    MessageBox.Show("Desconectou", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
    }
}

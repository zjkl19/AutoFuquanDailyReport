using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Diagnostics;
using System.Windows;
using AutoFuquanDailyReport.Services;

namespace AutoFuquanDailyReport
{
    public partial class MainWindow
    {
        private void OpenDataSummary_Click(object sender, RoutedEventArgs e)
        {


            string reportFile = $@"{FileService.GetFileName(App.OutputFolder, App.DataSummaryFile, string.Empty)}";
            if (File.Exists(reportFile))
            {
                Process.Start(reportFile);
            }
            else
            {
                MessageBox.Show($"请先生成数据汇总表。");
            }

        }
    }
}

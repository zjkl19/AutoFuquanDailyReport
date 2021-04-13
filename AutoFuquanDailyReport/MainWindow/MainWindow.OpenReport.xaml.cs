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
    public partial class MainWindow : Window
    {
        private void OpenReport_Click(object sender, RoutedEventArgs e)
        {
            
            //const string OutputReportFile = "自动生成的福泉互通病害整治工程--桥墩加固施工监测日报表.docx";

            string reportFile = $@"{FileService.GetFileName(App.OutputFolder, App.OutputReportFile, string.Empty)}";
            if (File.Exists(reportFile))
            {
                Process.Start(reportFile);
            }
            else
            {
                MessageBox.Show($"请先生成报告。");
            }

        }
    }
}

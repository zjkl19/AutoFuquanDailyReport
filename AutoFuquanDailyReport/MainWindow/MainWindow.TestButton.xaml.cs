using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Diagnostics;
using System.Windows;
using OfficeOpenXml;

namespace AutoFuquanDailyReport
{
    public partial class MainWindow : Window
    {
        private async void TestButton_Click(object sender, RoutedEventArgs e)
        {
            string info;
            FileInfo fileInfo = new FileInfo("20210409福泉主线日报数据处理.xlsx");
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(fileInfo))
            {
                var sheet = package.Workbook.Worksheets["桥墩及主梁实测数据实时录入（最后）"];
                //ExcelRange range = sheet.Cells[1, 1, 100000, 200];
                await package.SaveAsync();
                info = sheet.Cells[4, 11].Value.ToString();


            }
            MessageBox.Show(info);
        }
    }
}

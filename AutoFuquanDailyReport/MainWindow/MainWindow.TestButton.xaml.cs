using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Diagnostics;
using System.Windows;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.Style;
using AutoFuquanDailyReport.Services;
using System.Drawing.Imaging;
using System.Drawing;
using Spire.Xls;

namespace AutoFuquanDailyReport
{
    public partial class MainWindow : Window
    {
        private async void TestButton_Click(object sender, RoutedEventArgs e)
        {
            
            FileInfo fileInfo = new FileInfo(FileService.GetFileName(App.InputFolder, "数据汇总表", "xlsx"));
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(fileInfo))
            {
                var sheet = package.Workbook.Worksheets["桥墩水平位移Y"];

                var chart = sheet.Drawings["图表 1"] as ExcelLineChart;
                
                ExcelLineChartSerie series = chart.Series[0];
                for(int i=0;i<chart.Series.Count;i++)
                {
                    series = chart.Series[i];
                    series.Series = sheet.Cells[i+3, 4, i+3, 43].FullAddress;
                }
                FileInfo saveAsfileInfo = new FileInfo($"{App.InputFolder}\\Test数据汇总表.xlsx");
                await package.SaveAsAsync(saveAsfileInfo);

                //var excelPicture = sheet.Drawings[0] as ExcelPicture;
                
                //var img = excelPicture.Image; // This is of type System.Drawing.Image
                //img.Save(string.Format("img-1.png"), ImageFormat.Png);
            }
            //var workbook = new Workbook();
            //workbook.LoadFromFile(workbookFileName, ExcelVersion.Version2010);
            //var sheet = workbook.Worksheets[0]; // index or name of your worksheet
            //var image = workbook.SaveChartAsImage(sheet, 0); // chart index
            //img.Save(chartFileName, ImageFormat.Png);

            Workbook workbook = new Workbook();
            workbook.LoadFromFile(FileService.GetFileName(App.InputFolder, "数据汇总表", "xlsx"));
            Worksheet sheet1 = workbook.Worksheets["桥墩水平位移Y"];
            //遍历工作簿，诊断是否包含图表
            Image[] images = workbook.SaveChartAsImage(sheet1);

            for (int i = 0; i < images.Length; i++)
            {
                //将图表保存为图片
                images[i].Save(string.Format("img-{0}.png", i), ImageFormat.Png);
            }

            MessageBox.Show("测试完成");
        }
    }
}

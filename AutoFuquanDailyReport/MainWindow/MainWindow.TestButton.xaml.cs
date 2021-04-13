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
                //var sheet = package.Workbook.Worksheets["桥墩水平位移Y"];

                //var chart = sheet.Drawings["图表 1"] as ExcelLineChart;

                //ExcelLineChartSerie series = chart.Series[0];
                foreach(var sheet in package.Workbook.Worksheets)    //TODO：没有考虑没有工作表的情况
                {
                    if (sheet.Drawings.Count > 0)    //TODO：考虑有图片的情况
                    {
                        foreach (var drawing in sheet.Drawings)
                        {
                            var chart = drawing as ExcelLineChart;    //强制转换
                            if (chart.Series.Count > 0)
                            {
                                foreach (var series in chart.Series)
                                {
                                    var xAddr = new ExcelAddress(series.XSeries);    //x轴+1
                                    series.XSeries = sheet.Cells[xAddr.Start.Row, xAddr.Start.Column, xAddr.End.Row, xAddr.End.Column + 1].Address;

                                    var addr = new ExcelAddress(series.Series);    //数据量+1
                                    series.Series = sheet.Cells[addr.Start.Row, addr.Start.Column, addr.End.Row, addr.End.Column + 1].Address;
                                    //series.Series = sheet.Cells[i+3, 4, i+3, 43].Address;    //参考代码

                                }
                            }
                        }
                    }
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

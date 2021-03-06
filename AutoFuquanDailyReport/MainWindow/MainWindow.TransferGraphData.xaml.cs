using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Diagnostics;
using System.Windows;
using OfficeOpenXml;
using AutoFuquanDailyReport.Services;
using OfficeOpenXml.Drawing.Chart;

namespace AutoFuquanDailyReport
{
    public partial class MainWindow : Window
    {
        //将“福泉主线日报数据处理”中的数据逐个复制到“数据汇总表”中
        private async void TransferGraphData_Click(object sender, RoutedEventArgs e)
        {
            string info;

            //var k = FileService.GetFileName(App.InputFolder, "福泉主线日报数据处理","xlsx");
            //FileInfo fileInfo = new FileInfo($"{App.InputFolder}\\福泉主线日报数据处理.xlsx");
            FileInfo fileInfo = new FileInfo(FileService.GetFileName(App.InputFolder, "福泉主线日报数据处理", "xlsx"));
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var package = new ExcelPackage(fileInfo);

            var sheet = package.Workbook.Worksheets["桥墩及主梁实测数据实时录入（最后）"];
            var sheetOfPierAndBeam = package.Workbook.Worksheets["桥墩及主梁报告内表格数据"];
            var sheetOfGround = package.Workbook.Worksheets["地表沉降报告内数据"];

            //从原始数据表读取数据
            //桥墩
            const int PierNodes = 35; const int PierRowIndex = 4;
            decimal[,] PierData = new decimal[35, 3];
            for (int i = 0; i < PierNodes; i++)
            {
                PierData[i, 0] = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + PierRowIndex, 6].Value);    //Y
                PierData[i, 1] = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + PierRowIndex, 7].Value);    //X
                PierData[i, 2] = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + PierRowIndex, 8].Value);    //Z
            }
            const int PerpNodes = 15; const int PerpRowIndex = 61;
            decimal[,] PerpData = new decimal[PerpNodes, 3];
            for (int i = 0; i < PerpNodes; i++)
            {
                PerpData[i, 0] = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + PerpRowIndex, 5].Value);
                PerpData[i, 1] = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + PerpRowIndex, 6].Value);
            }

            //地表
            const int GroundNodes = 18; const int GroundRowIndex = 4;
            decimal[,] groundData = new decimal[GroundNodes, 1];
            int[] GroundRowIndexList = { 4, 5, 6, 7, 8, 9, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33 };

            for (int i = 0; i < GroundNodes; i++)
            {
                try
                {
                    groundData[i, 0] = Convert.ToDecimal(sheetOfGround.Cells[GroundRowIndexList[i], 9].Value);
                }
                catch (Exception)
                {

                    groundData[i, 0] = App.AbnormalData;
                }
            }
            //涵洞
            const int CulvertNodes = 28; const int CulvertRowIndex = 34;
            decimal[,] culvertData = new decimal[CulvertNodes, 1];
            for (int i = 0; i < CulvertNodes; i++)
            {
                try
                {
                    culvertData[i, 0] = Convert.ToDecimal(sheetOfGround.Cells[i + CulvertRowIndex, 9].Value);
                }
                catch (Exception)
                {
                    culvertData[i, 0] = App.AbnormalData;
                }
            }

            //主梁
            const int BeamNodes = 15; const int BeamRowIndex = 40;
            decimal[,] BeamData = new decimal[BeamNodes, 12];
            for (int i = 0; i < BeamNodes; i++)
            {
                BeamData[i, 0] = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + BeamRowIndex, 6].Value);
                BeamData[i, 1] = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + BeamRowIndex, 7].Value);
                BeamData[i, 2] = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + BeamRowIndex, 8].Value);
            }

            //将数据写入到汇总表中

            //FileInfo saveFileInfo = new FileInfo($"{App.InputFolder}\\数据汇总表.xlsx");    //参考代码
            FileInfo saveFileInfo = new FileInfo(FileService.GetFileName(App.InputFolder, "数据汇总表", "xlsx"));

            var savePackage = new ExcelPackage(saveFileInfo);
            var sheetOfPierY = savePackage.Workbook.Worksheets["桥墩水平位移Y"];
            var sheetOfPierX = savePackage.Workbook.Worksheets["桥墩水平位移X"];
            var sheetOfPierZ = savePackage.Workbook.Worksheets["桥墩沉降Z"];

            var sheetOfPerpY = savePackage.Workbook.Worksheets["桥墩垂直度Y"];
            var sheetOfPerpX = savePackage.Workbook.Worksheets["桥墩垂直度X"];

            var sheetOfGroundZ = savePackage.Workbook.Worksheets["地表及路基沉降"];
            var sheetOfCulvertZ = savePackage.Workbook.Worksheets["涵洞沉降"];

            var sheetOfBeamY = savePackage.Workbook.Worksheets["主梁Y"];
            var sheetOfBeamX = savePackage.Workbook.Worksheets["主梁X"];
            var sheetOfBeamZ = savePackage.Workbook.Worksheets["主梁Z"];

            const int MaxSearchCol = 3000;    //最大搜索列数
            const int SavePierRowIndex = 2;
            int colCurr = 1;
            //查找日期的空行
            colCurr = SearchCol(sheetOfPierY, MaxSearchCol, SavePierRowIndex, 1);
            //日期设置
            string dt = sheetOfPierY.Cells[SavePierRowIndex, colCurr - 1].Value.ToString();
            DateTime day;
            System.Globalization.DateTimeFormatInfo dtFormat = new System.Globalization.DateTimeFormatInfo
            {
                ShortDatePattern = "yyyy.MM.dd"
            };
            day = Convert.ToDateTime(dt, dtFormat);
            string dateInWorksheet = day.AddDays(1).ToString("yyyy.MM.dd");

            sheetOfPierY.Cells[SavePierRowIndex, colCurr].Value = dateInWorksheet;
            for (int i = 0; i < PierNodes; i++)
            {
                sheetOfPierY.Cells[i + SavePierRowIndex + 1, colCurr].Value = Math.Round(PierData[i, 0], 1);
            }

            colCurr = SearchCol(sheetOfPierX, MaxSearchCol, SavePierRowIndex, 2);
            sheetOfPierX.Cells[SavePierRowIndex, colCurr].Value = dateInWorksheet;
            for (int i = 0; i < PierNodes; i++)
            {
                sheetOfPierX.Cells[i + SavePierRowIndex + 1, colCurr].Value = Math.Round(PierData[i, 1], 1);
            }

            colCurr = SearchCol(sheetOfPierZ, MaxSearchCol, SavePierRowIndex, 2);
            sheetOfPierZ.Cells[SavePierRowIndex, colCurr].Value = dateInWorksheet;
            for (int i = 0; i < PierNodes; i++)
            {
                sheetOfPierZ.Cells[i + SavePierRowIndex + 1, colCurr].Value = Math.Round(PierData[i, 2], 1);
            }

            const int SavePerpRowIndex = 2;
            colCurr = SearchCol(sheetOfPerpY, MaxSearchCol, SavePerpRowIndex, 2);
            sheetOfPerpY.Cells[SavePerpRowIndex, colCurr].Value = dateInWorksheet;
            for (int i = 0; i < PerpNodes; i++)
            {
                sheetOfPerpY.Cells[i + SavePerpRowIndex + 1, colCurr].Value = PerpData[i, 0];
                sheetOfPerpY.Cells[i + SavePerpRowIndex + 1, colCurr].Style.Numberformat.Format = "0.00%";
            }


            colCurr = SearchCol(sheetOfPerpX, MaxSearchCol, SavePerpRowIndex, 2);
            sheetOfPerpX.Cells[SavePerpRowIndex, colCurr].Value = dateInWorksheet;
            for (int i = 0; i < PerpNodes; i++)
            {
                sheetOfPerpX.Cells[i + SavePerpRowIndex + 1, colCurr].Value = PerpData[i, 1]; //Math.Round(PerpData[i, 1], 1);
                sheetOfPerpX.Cells[i + SavePerpRowIndex + 1, colCurr].Style.Numberformat.Format = "0.00%";
            }

            const int DefaultSaveRowIndex = 2;
            colCurr = SearchCol(sheetOfGroundZ, MaxSearchCol, DefaultSaveRowIndex, 3);
            sheetOfGroundZ.Cells[DefaultSaveRowIndex, colCurr].Value = dateInWorksheet;
            for (int i = 0; i < GroundNodes; i++)
            {
                sheetOfGroundZ.Cells[i + DefaultSaveRowIndex + 1, colCurr].Value = Math.Round(groundData[i, 0], 1);
            }

            colCurr = SearchCol(sheetOfCulvertZ, MaxSearchCol, DefaultSaveRowIndex, 3);
            sheetOfCulvertZ.Cells[DefaultSaveRowIndex, colCurr].Value = dateInWorksheet;
            for (int i = 0; i < CulvertNodes; i++)
            {
                if (Math.Abs(culvertData[i, 0]) > 100)
                {
                    sheetOfCulvertZ.Cells[i + DefaultSaveRowIndex + 1, colCurr].Value = "/";
                }
                else
                {
                    sheetOfCulvertZ.Cells[i + DefaultSaveRowIndex + 1, colCurr].Value = Math.Round(culvertData[i, 0], 1);
                }
            }

            colCurr = SearchCol(sheetOfBeamY, MaxSearchCol, DefaultSaveRowIndex, 3);
            sheetOfBeamY.Cells[DefaultSaveRowIndex, colCurr].Value = dateInWorksheet;
            for (int i = 0; i < BeamNodes; i++)
            {
                sheetOfBeamY.Cells[i + DefaultSaveRowIndex + 1, colCurr].Value = Math.Round(BeamData[i, 0], 1);
            }

            colCurr = SearchCol(sheetOfBeamX, MaxSearchCol, DefaultSaveRowIndex, 3);
            sheetOfBeamX.Cells[DefaultSaveRowIndex, colCurr].Value = dateInWorksheet;
            for (int i = 0; i < BeamNodes; i++)
            {
                sheetOfBeamX.Cells[i + DefaultSaveRowIndex + 1, colCurr].Value = Math.Round(BeamData[i, 1], 1);
            }

            colCurr = SearchCol(sheetOfBeamZ, MaxSearchCol, DefaultSaveRowIndex, 3);
            sheetOfBeamZ.Cells[DefaultSaveRowIndex, colCurr].Value = dateInWorksheet;
            for (int i = 0; i < BeamNodes; i++)
            {
                sheetOfBeamZ.Cells[i + DefaultSaveRowIndex + 1, colCurr].Value = Math.Round(BeamData[i, 2], 1);
            }


            //var sheet = package.Workbook.Worksheets["桥墩水平位移Y"];

            //var chart = sheet.Drawings["图表 1"] as ExcelLineChart;

            //ExcelLineChartSerie series = chart.Series[0];
            UpdateExcelChart(savePackage);

            FileInfo saveAsFileInfo = new FileInfo($"{App.OutputFolder}\\{day.AddDays(1):yyyyMMdd}数据汇总表.xlsx");

            // Save our new workbook in the output directory and we are done!
            await savePackage.SaveAsAsync(saveAsFileInfo);

            MessageBox.Show("数据复制完成！");
        }
        /// <summary>
        /// 保存的excel表格中Chart数据更新
        /// </summary>
        /// <param name="savePackage"></param>
        private static void UpdateExcelChart(ExcelPackage savePackage)
        {
            foreach (var sheet in savePackage.Workbook.Worksheets)    //TODO：没有考虑没有工作表的情况
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
        }

        /// <summary>
        /// 查找第1个空列
        /// </summary>
        /// <param name="sheetOfPierY">ExcelWorksheet</param>
        /// <param name="MaxSearchCol">最大查找列</param>
        /// <param name="SavePierRowIndex">从第几行开始查找</param>
        /// <param name="colCurr">从第几列开始查找</param>
        /// <returns></returns>
        private static int SearchCol(ExcelWorksheet sheetOfPierY, int MaxSearchCol, int SavePierRowIndex, int colCurr)
        {
            while (colCurr < MaxSearchCol)
            {
                if (string.IsNullOrWhiteSpace(sheetOfPierY.Cells[SavePierRowIndex, colCurr]?.Value?.ToString() ?? string.Empty))
                {
                    break;
                }
                colCurr++;
            }

            return colCurr;
        }
    }
}

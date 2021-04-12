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
        //将“福泉主线日报数据处理”中的数据逐个复制到“数据汇总表”中
        private async void TransferGraphData_Click(object sender, RoutedEventArgs e)
        {
            string info;
            FileInfo fileInfo = new FileInfo($"{App.InputFolder}\\福泉主线日报数据处理.xlsx");
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

            //将数据写入到汇总表中
            FileInfo saveFileInfo = new FileInfo($"{App.InputFolder}\\数据汇总表.xlsx");

            var savePackage = new ExcelPackage(saveFileInfo);
            var sheetOfPierY = savePackage.Workbook.Worksheets["桥墩水平位移Y"];
            var sheetOfPierX = savePackage.Workbook.Worksheets["桥墩水平位移X"];
            var sheetOfPierZ = savePackage.Workbook.Worksheets["桥墩沉降Z"];

            var sheetOfPerpY = savePackage.Workbook.Worksheets["桥墩垂直度Y"];
            var sheetOfPerpX = savePackage.Workbook.Worksheets["桥墩垂直度X"];

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
                sheetOfPerpY.Cells[i + SavePerpRowIndex + 1, colCurr].Value = $"{PerpData[i, 0]:P}";
            }

            
            colCurr = SearchCol(sheetOfPerpX, MaxSearchCol, SavePerpRowIndex, 2);
            sheetOfPerpX.Cells[SavePerpRowIndex, colCurr].Value = dateInWorksheet;
            for (int i = 0; i < PerpNodes; i++)
            {
                sheetOfPerpX.Cells[i + SavePerpRowIndex + 1, colCurr].Value = $"{PerpData[i, 1]:P}"; //Math.Round(PerpData[i, 1], 1);
            }


            FileInfo saveAsFileInfo = new FileInfo($"{App.OutputFolder}\\{day.AddDays(1):yyyyMMdd}数据汇总表.xlsx");

            // Save our new workbook in the output directory and we are done!
            await savePackage.SaveAsAsync(saveAsFileInfo);

            MessageBox.Show("数据复制完成！");
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

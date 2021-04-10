﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Aspose.Words;
using Aspose.Words.Tables;
using OfficeOpenXml;

namespace AutoFuquanDailyReport
{
    public partial class MainWindow : Window
    {
        private async void AutoReport_Click(object sender, RoutedEventArgs e)
        {
            string templateFile = @"Templates\福泉互通病害整治工程--桥墩加固施工监测日报表模板.docx";
            string outputFile = @"OutputReport\自动生成的福泉互通病害整治工程--桥墩加固施工监测日报表.docx";

            string info;
            FileInfo fileInfo = new FileInfo(@"OriginalData\福泉主线日报数据处理.xlsx");
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var package = new ExcelPackage(fileInfo);

            var sheetOfPierAndBeam = package.Workbook.Worksheets["桥墩及主梁报告内表格数据"];
            //ExcelRange range = sheet.Cells[1, 1, 100000, 200];
            await package.SaveAsync();
            info = sheetOfPierAndBeam.Cells[4, 11].Value.ToString();

            const int dataRows = 24; const int dataColumns = 8;

            //东、西主桥、D匝道桥墩测点水平位移、沉降监测数据汇总表
            decimal[,] PierData = new decimal[35, 12];
            for (int i = 0; i < 35; i++)
            {

                //前次累积值
                PierData[i, 0] = Convert.ToDecimal(sheetOfPierAndBeam.Cells[4 + i, 11].Value);
                PierData[i, 1] = Convert.ToDecimal(sheetOfPierAndBeam.Cells[4 + i, 12].Value);
                PierData[i, 2] = Convert.ToDecimal(sheetOfPierAndBeam.Cells[4 + i, 13].Value);

                //本次累积值
                PierData[i, 3] = Convert.ToDecimal(sheetOfPierAndBeam.Cells[4 + i, 6].Value);
                PierData[i, 4] = Convert.ToDecimal(sheetOfPierAndBeam.Cells[4 + i, 7].Value);
                PierData[i, 5] = Convert.ToDecimal(sheetOfPierAndBeam.Cells[4 + i, 8].Value);

                //本次变化值
                PierData[i, 6] = Convert.ToDecimal(sheetOfPierAndBeam.Cells[4 + i, 3].Value);
                PierData[i, 7] = Convert.ToDecimal(sheetOfPierAndBeam.Cells[4 + i, 4].Value);
                PierData[i, 8] = Convert.ToDecimal(sheetOfPierAndBeam.Cells[4 + i, 5].Value);

                //本次变化速率
                PierData[i, 9] = Convert.ToDecimal(sheetOfPierAndBeam.Cells[4 + i, 3].Value);
                PierData[i, 10] = Convert.ToDecimal(sheetOfPierAndBeam.Cells[4 + i, 4].Value);
                PierData[i, 11] = Convert.ToDecimal(sheetOfPierAndBeam.Cells[4 + i, 5].Value);

            }

            //东、西主桥、D匝道桥墩垂直度监测数据汇总表
            const int PerpRowIndex = 61;
            decimal[,] PerpData = new decimal[15, 8];
            for (int i = 0; i < 15; i++)
            {
                //前次累积值
                PerpData[i, 0] = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + PerpRowIndex, 11].Value);
                PerpData[i, 1] = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + PerpRowIndex, 12].Value);

                //本次累积值
                PerpData[i, 2] = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + PerpRowIndex, 5].Value);
                PerpData[i, 3] = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + PerpRowIndex, 6].Value);

                //本次变化值
                PerpData[i, 4] = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + PerpRowIndex, 3].Value);
                PerpData[i, 5] = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + PerpRowIndex, 4].Value);

                //本次变化速率
                PerpData[i, 6] = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + PerpRowIndex, 3].Value);
                PerpData[i, 7] = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + PerpRowIndex, 4].Value);
            }

            //东、西主桥、D匝道主梁测点水平位移、沉降监测数据汇总表
            const int BeamNodes = 15; const int BeamRowIndex = 40;
            decimal[,] BeamData = new decimal[BeamNodes, 12];
            for (int i = 0; i < BeamNodes; i++)
            {
                //前次累积值
                BeamData[i, 0] = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + BeamRowIndex, 11].Value);
                BeamData[i, 1] = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + BeamRowIndex, 12].Value);
                BeamData[i, 2] = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + BeamRowIndex, 13].Value);

                //本次累积值
                BeamData[i, 3] = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + BeamRowIndex, 6].Value);
                BeamData[i, 4] = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + BeamRowIndex, 7].Value);
                BeamData[i, 5] = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + BeamRowIndex, 8].Value);
                //本次变化值
                BeamData[i, 6] = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + BeamRowIndex, 3].Value);
                BeamData[i, 7] = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + BeamRowIndex, 4].Value);
                BeamData[i, 8] = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + BeamRowIndex, 5].Value);

                //本次变化速率
                BeamData[i, 9] = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + BeamRowIndex, 3].Value);
                BeamData[i, 10] = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + BeamRowIndex, 4].Value);
                BeamData[i, 11] = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + BeamRowIndex, 5].Value);

            }
            //MessageBox.Show(info);

            try
            {
                var doc = new Document(templateFile);


                var builder = new DocumentBuilder(doc);

                Table table0 = doc.GetChildNodes(NodeType.Table, true)[0] as Table;

                Table perpTable = doc.GetChildNodes(NodeType.Table, true)[2] as Table;    //东、西主桥、D匝道桥墩垂直度监测数据汇总表

                Table beamTable = doc.GetChildNodes(NodeType.Table, true)[3] as Table;    //东、西主桥、D匝道主梁测点水平位移、沉降监测数据汇总表

                //东、西主桥、D匝道桥墩测点水平位移、沉降监测数据汇总表
                for (int i = 0; i < PierData.GetLength(0); i++)
                {
                    for (int j = 0; j < PierData.GetLength(1); j++)
                    {
                        builder.MoveTo(table0.Rows[i + 2].Cells[j + 2].FirstParagraph);    //3行2列
                        builder.Write(PierData[i, j].ToString());
                    }

                }
                //东、西主桥、D匝道桥墩垂直度监测数据汇总表
                for (int i = 0; i < PerpData.GetLength(0); i++)
                {
                    for (int j = 0; j < PerpData.GetLength(1); j++)
                    {
                        builder.MoveTo(perpTable.Rows[i + 2].Cells[j + 1].FirstParagraph);    //3行2列
                        builder.Write($"{PerpData[i, j]:P}");
                    }

                }

                //东、西主桥、D匝道主梁测点水平位移、沉降监测数据汇总表
                for (int i = 0; i < BeamData.GetLength(0); i++)
                {
                    for (int j = 0; j < BeamData.GetLength(1); j++)
                    {
                        builder.MoveTo(beamTable.Rows[i + 2].Cells[j + 1].FirstParagraph);    //3行2列
                        builder.Write($"{BeamData[i, j]}");
                    }

                }

                //excel表格"桥墩及主梁报告内表格数据"标签中，前一日累计变化值(mm)、累积变化值(mm)及本次变化值(mm)
                //依次复制到Word相应表格中


                doc.UpdateFields();
                doc.Save(outputFile, SaveFormat.Docx);
                MessageBox.Show("成功生成报告！");

            }
            catch (Exception ex)
            {

                Debug.Print("报告生成失败，请检查原因！");
            }
        }
    }
}

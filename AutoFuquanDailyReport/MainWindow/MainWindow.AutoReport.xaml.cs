using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Aspose.Words;
using Aspose.Words.Fields;
using Aspose.Words.Tables;
using AutoFuquanDailyReport.Models;
using AutoFuquanDailyReport.Services;
using OfficeOpenXml;

namespace AutoFuquanDailyReport
{
    public partial class MainWindow : Window
    {
        private async void AutoReport_Click(object sender, RoutedEventArgs e)
        {
            string templateFile = $"{App.TemplateFolder}\\福泉互通病害整治工程--桥墩加固施工监测日报表模板.docx";
            string outputFile = $"{App.OutputFolder}\\{DateTime.Now:yyyyMMdd}自动生成的福泉互通病害整治工程--桥墩加固施工监测日报表.docx";

            string info;

            //FileInfo fileInfo = new FileInfo($"{App.InputFolder}\\福泉主线日报数据处理.xlsx");
            FileInfo fileInfo = new FileInfo(FileService.GetFileName(App.InputFolder, "福泉主线日报数据处理", "xlsx"));
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var package = new ExcelPackage(fileInfo);

            var sheetOfPierAndBeam = package.Workbook.Worksheets["桥墩及主梁报告内表格数据"];
            var sheetOfGround = package.Workbook.Worksheets["地表沉降报告内数据"];

            //ExcelRange range = sheet.Cells[1, 1, 100000, 200];
            await package.SaveAsync();
            info = sheetOfPierAndBeam.Cells[4, 11].Value.ToString();

            const int dataRows = 24; const int dataColumns = 8;

            //东、西主桥、D匝道桥墩测点水平位移、沉降监测数据汇总表
            decimal[,] PierData = new decimal[35, 12];
            //数据合并成一个表格，方便插入word
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

            //List<PierMeasurePoint> pierMeasurePoints = new List<PierMeasurePoint>(App.PierMeasurePointCounts);

            PierMeasurePoint[] pierMeasurePoints = new PierMeasurePoint[App.PierMeasurePointCounts];

            for (int i = 0; i < App.PierMeasurePointCounts; i++)
            {
                pierMeasurePoints[i] = new PierMeasurePoint
                {
                    No = sheetOfPierAndBeam.Cells[4 + i, 2]?.Value?.ToString() ?? string.Empty
                    ,
                    //前次累积值
                    PreviousAccumulateY = Convert.ToDecimal(sheetOfPierAndBeam.Cells[4 + i, 11].Value)
                    ,
                    PreviousAccumulateX = Convert.ToDecimal(sheetOfPierAndBeam.Cells[4 + i, 12].Value)
                    ,
                    PreviousAccumulateZ = Convert.ToDecimal(sheetOfPierAndBeam.Cells[4 + i, 13].Value)
                    //本次累积值
                    ,
                    AccumulateY = Convert.ToDecimal(sheetOfPierAndBeam.Cells[4 + i, 6].Value)
                    ,
                    AccumulateX = Convert.ToDecimal(sheetOfPierAndBeam.Cells[4 + i, 7].Value)
                    ,
                    AccumulateZ = Convert.ToDecimal(sheetOfPierAndBeam.Cells[4 + i, 8].Value)
                    //本次变化值
                    ,
                    DeltaY = Convert.ToDecimal(sheetOfPierAndBeam.Cells[4 + i, 3].Value)
                    ,
                    DeltaX = Convert.ToDecimal(sheetOfPierAndBeam.Cells[4 + i, 4].Value)
                    ,
                    DeltaZ = Convert.ToDecimal(sheetOfPierAndBeam.Cells[4 + i, 5].Value)

                    //本次变化速率
                    ,
                    RateOfChangeY = Convert.ToDecimal(sheetOfPierAndBeam.Cells[4 + i, 3].Value)
                    ,
                    RateOfChangeX = Convert.ToDecimal(sheetOfPierAndBeam.Cells[4 + i, 4].Value)
                    ,
                    RateOfChangeZ = Convert.ToDecimal(sheetOfPierAndBeam.Cells[4 + i, 5].Value)
                };
                //TODO：消除硬编码（放到excel中）
                pierMeasurePoints[i].YDirection = pierMeasurePoints[i].GetYDirection("东", "西");    //每个测点要分开设置，以下同理
                pierMeasurePoints[i].XDirection = pierMeasurePoints[i].GetXDirection("大里程", "小里程");
                pierMeasurePoints[i].ZDirection = pierMeasurePoints[i].GetXDirection("上", "下");    //20210417暂定
            }

            //东、西主桥、D匝道桥墩垂直度监测数据汇总表
            const int PerpRowIndex = 61;
            decimal[,] PerpData = new decimal[15, 8];

            PerpMeasurePoint[] perpMeasurePoints = new PerpMeasurePoint[App.PerpMeasurePointCounts];
            for (int i = 0; i < App.PerpMeasurePointCounts; i++)
            {
                perpMeasurePoints[i] = new PerpMeasurePoint
                {
                    No = sheetOfPierAndBeam.Cells[PerpRowIndex + i, 2]?.Value?.ToString() ?? string.Empty
                    ,
                    //前次累积值
                    PreviousAccumulateY = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + PerpRowIndex, 11].Value)
                    ,
                    PreviousAccumulateX = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + PerpRowIndex, 12].Value)
                    //本次累积值
                    ,
                    AccumulateY = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + PerpRowIndex, 5].Value)
                    ,
                    AccumulateX = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + PerpRowIndex, 6].Value)
                    //本次变化值
                    ,
                    DeltaY = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + PerpRowIndex, 3].Value)
                    ,
                    DeltaX = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + PerpRowIndex, 4].Value)

                    //本次变化速率
                    ,
                    RateOfChangeY = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + PerpRowIndex, 3].Value)
                    ,
                    RateOfChangeX = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + PerpRowIndex, 4].Value)
                };
            }

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
            var beamMeasurePoints = new BeamMeasurePoint[App.BeamMeasurePointCounts];
            for (int i = 0; i < App.BeamMeasurePointCounts; i++)
            {
                beamMeasurePoints[i] = new BeamMeasurePoint
                {
                    No = sheetOfPierAndBeam.Cells[i + BeamRowIndex, 2]?.Value?.ToString() ?? string.Empty,
                    //前次累积值
                    PreviousAccumulateY = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + BeamRowIndex, 11].Value),
                    PreviousAccumulateX = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + BeamRowIndex, 12].Value),
                    PreviousAccumulateZ = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + BeamRowIndex, 13].Value),

                    //本次累积值
                    AccumulateY = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + BeamRowIndex, 6].Value),
                    AccumulateX = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + BeamRowIndex, 7].Value),
                    AccumulateZ = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + BeamRowIndex, 8].Value),
                    //本次变化值
                    DeltaY = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + BeamRowIndex, 3].Value),
                    DeltaX = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + BeamRowIndex, 4].Value),
                    DeltaZ = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + BeamRowIndex, 5].Value),

                    //本次变化速率
                    RateOfChangeY = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + BeamRowIndex, 3].Value),
                    RateOfChangeX = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + BeamRowIndex, 4].Value),
                    RateOfChangeZ = Convert.ToDecimal(sheetOfPierAndBeam.Cells[i + BeamRowIndex, 5].Value),

                };
                //TODO：消除硬编码（放到excel中）
                beamMeasurePoints[i].YDirection = beamMeasurePoints[i].GetYDirection("东", "西");    //每个测点要分开设置，以下同理
                beamMeasurePoints[i].XDirection = beamMeasurePoints[i].GetXDirection("大里程", "小里程");
                beamMeasurePoints[i].ZDirection = beamMeasurePoints[i].GetXDirection("上", "下");    //20210417暂定

            }

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


            //涵洞测点沉降监测数据汇总表
            const int GroundNodes = 18; const int GroundRowIndex = 4;
            decimal[,] groundData = new decimal[GroundNodes, 4];
            int[] GroundRowIndexList = { 4, 5, 6, 7, 8, 9, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33 };
            int[] tIndex = { 14, 9, 6, 6 };    //excel中所在列
            int k;

            for (int i = 0; i < GroundNodes; i++)
            {
                k = 0;
                foreach (var item in tIndex)
                {
                    try
                    {
                        groundData[i, k] = Convert.ToDecimal(sheetOfGround.Cells[GroundRowIndexList[i], item].Value);
                    }
                    catch (Exception)
                    {

                        groundData[i, k] = -9999m;
                    }

                    k++;
                }

                //前次累积值
                //groundData[i, 0] = Convert.ToDecimal(sheetOfGround.Cells[i + GroundRowIndex, 14].Value);
                //本次累积值
                //groundData[i, 1] = Convert.ToDecimal(sheetOfGround.Cells[i + GroundRowIndex, 9].Value);
                //本次变化值
                //groundData[i, 2] = Convert.ToDecimal(sheetOfGround.Cells[i + GroundRowIndex, 6].Value);
                //本次变化速率
                //groundData[i, 3] = Convert.ToDecimal(sheetOfGround.Cells[i + GroundRowIndex, 6].Value);
            }

            const int CulvertNodes = 28; const int CulvertRowIndex = 34;
            decimal[,] culvertData = new decimal[CulvertNodes, 4];
            for (int i = 0; i < CulvertNodes; i++)
            {
                k = 0;
                foreach (var item in tIndex)
                {
                    try
                    {
                        culvertData[i, k] = Convert.ToDecimal(sheetOfGround.Cells[i + CulvertRowIndex, item].Value);
                    }
                    catch (Exception)
                    {

                        culvertData[i, k] = -9999m;
                    }

                    k++;
                }
            }

            try
            {
                string[] MyDocumentVariables = new string[] { "ReportDate", "MonitorTime", "PierAndPerpConclusion","BeamConclusion" };//文档中包含的所有“文档变量”，方便遍历

                var maxYPierMeasurePoints = pierMeasurePoints.OrderByDescending(s => Math.Abs(s.AccumulateY)).First();
                var maxXPierMeasurePoints = pierMeasurePoints.OrderByDescending(s => Math.Abs(s.AccumulateX)).First();
                var maxZPierMeasurePoints = pierMeasurePoints.OrderBy(s => s.AccumulateZ).First();    //实际上是最小值

                var minYPerpMeasurePoints = perpMeasurePoints.OrderBy(s => s.AccumulateY).First();
                var minXPerpMeasurePoints = perpMeasurePoints.OrderBy(s => s.AccumulateX).First();

                var maxYPerpMeasurePoints = perpMeasurePoints.OrderByDescending(s => s.AccumulateY).First();
                var maxXPerpMeasurePoints = perpMeasurePoints.OrderByDescending(s => s.AccumulateX).First();

                string pierConclusion = $"桥墩横桥向累计最大水平位移为向{maxYPierMeasurePoints.YDirection}侧{Math.Abs(maxYPierMeasurePoints.AccumulateY):F1}mm（{maxYPierMeasurePoints.No}测点）；纵桥向累计最大水平位移为向{maxXPierMeasurePoints.XDirection}{Math.Abs(maxXPierMeasurePoints.AccumulateX):F1}mm（{maxXPierMeasurePoints.No}测点），最大沉降为{-1.0m*(maxZPierMeasurePoints.AccumulateZ):F1}mm（{maxZPierMeasurePoints.No}测点）。桥墩垂直度累计变化值在{Math.Min(minYPerpMeasurePoints.AccumulateY, minXPerpMeasurePoints.AccumulateX):P}~{Math.Max(maxYPerpMeasurePoints.AccumulateY, maxXPerpMeasurePoints.AccumulateX):P}之间。";
                //注意垂直度把x和y放在一起比较的方法可能有问题

                var maxYBeamMeasurePoints = beamMeasurePoints.OrderByDescending(s => Math.Abs(s.AccumulateY)).First();
                var maxXBeamMeasurePoints = beamMeasurePoints.OrderByDescending(s => Math.Abs(s.AccumulateX)).First();
                var maxZBeamMeasurePoints = beamMeasurePoints.OrderBy(s => s.AccumulateZ).First();    //实际上是最小值

                string beamConclusion = $"主梁测点横桥向累计最大水平位移为向{maxYBeamMeasurePoints.YDirection}侧{Math.Abs(maxYBeamMeasurePoints.AccumulateY):F1}mm（{maxYBeamMeasurePoints.No}主梁测点）；纵桥向累计最大水平位移为向{maxXBeamMeasurePoints.XDirection}{Math.Abs(maxXBeamMeasurePoints.AccumulateX)}mm（{maxXBeamMeasurePoints.No}主梁测点），最大沉降为{-1.0m*(maxZBeamMeasurePoints.AccumulateZ):F1}mm（{maxZBeamMeasurePoints.No}主梁测点）。";

                var doc = new Document(templateFile);

                //更新文档变量
                try
                {
                    var variables = doc.Variables;
                    variables["ReportDate"] = (ReportTime.SelectedDate ?? DateTime.Now).ToString("yyyy年MM月dd日");
                    variables["MonitorTime"] = (ReportTime.SelectedDate ?? DateTime.Now).ToString("yyyy年MM月dd日");
                    variables["PierAndPerpConclusion"] = pierConclusion;
                    variables["BeamConclusion"] = beamConclusion;
                }
                catch (Exception ex)
                {

                    Debug.Print($"文档变量更新失败。信息{ex.Message}");
                }

                doc.UpdateFields();

                foreach (var v in doc.Range.Fields)
                {
                    var v1 = (FieldDocVariable)v;
                    if (MyDocumentVariables.Contains(v1.VariableName))
                    {
                        v1.Unlink();
                    }

                }

                var builder = new DocumentBuilder(doc);

                Table table0 = doc.GetChildNodes(NodeType.Table, true)[0] as Table;

                Table perpTable = doc.GetChildNodes(NodeType.Table, true)[2] as Table;    //东、西主桥、D匝道桥墩垂直度监测数据汇总表

                Table beamTable = doc.GetChildNodes(NodeType.Table, true)[3] as Table;    //东、西主桥、D匝道主梁测点水平位移、沉降监测数据汇总表

                Table groundTable = doc.GetChildNodes(NodeType.Table, true)[4] as Table;    //地表沉降及路基横断面测点沉降监测数据汇总表

                Table culvertTable = doc.GetChildNodes(NodeType.Table, true)[5] as Table;    //涵洞测点沉降监测数据汇总表



                //东、西主桥、D匝道桥墩测点水平位移、沉降监测数据汇总表
                for (int i = 0; i < PierData.GetLength(0); i++)
                {
                    for (int j = 0; j < PierData.GetLength(1); j++)
                    {
                        builder.MoveTo(table0.Rows[i + 2].Cells[j + 2].FirstParagraph);    //3行2列
                        builder.Write($"{PierData[i, j]:F1}");
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
                        builder.Write($"{BeamData[i, j]:F1}");
                    }
                }
                //地表沉降及路基横断面测点沉降监测数据汇总表
                for (int i = 0; i < groundData.GetLength(0); i++)
                {
                    for (int j = 0; j < groundData.GetLength(1); j++)
                    {
                        builder.MoveTo(groundTable.Rows[i + 1].Cells[j + 1].FirstParagraph);    //3行2列
                        builder.Write($"{groundData[i, j]:F1}");
                    }
                }

                //涵洞测点沉降监测数据汇总表
                for (int i = 0; i < culvertData.GetLength(0); i++)
                {
                    for (int j = 0; j < culvertData.GetLength(1); j++)
                    {
                        builder.MoveTo(culvertTable.Rows[i + 1].Cells[j + 1].FirstParagraph);    //3行2列
                        if (Math.Abs(culvertData[i, j]) > 100)
                        {
                            builder.Write($"/");
                        }
                        else
                        {
                            builder.Write($"{culvertData[i, j]:F1}");
                        }

                    }
                }

                //excel表格"桥墩及主梁报告内表格数据"标签中，前一日累计变化值(mm)、累积变化值(mm)及本次变化值(mm)
                //依次复制到Word相应表格中


                doc.UpdateFields();
                doc.Save(outputFile, SaveFormat.Docx);
                MessageBox.Show("成功生成报告！请再核对一遍自动报告的结果，防止出错。");

            }
            catch (Exception ex)
            {

                Debug.Print("报告生成失败，请检查原因！");
            }
        }
    }
}

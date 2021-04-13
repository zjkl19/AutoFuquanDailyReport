using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace AutoFuquanDailyReport
{
    /// <summary>
    /// App.xaml 的交互逻辑
    /// </summary>
    public partial class App : Application
    {
        //全局常量

        //参考https://www.cnblogs.com/Gildor/archive/2010/06/29/1767156.html
        public static double ScreenWidth = SystemParameters.PrimaryScreenWidth;
        public static double ScreenHeight = SystemParameters.PrimaryScreenHeight;

        public static string InputFolder = "Input";
        public static string OutputFolder = "Output";
        public static string TemplateFolder = "Templates";
        public static string OutputReportFile = "自动生成的福泉互通病害整治工程--桥墩加固施工监测日报表.docx";


        public static decimal NotApplicableData = 200m;    //仅为正值，若数据绝对值大于该值，则值为异常值
        public static decimal AbnormalData = -9999m;
    }
}

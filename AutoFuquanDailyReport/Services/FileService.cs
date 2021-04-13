using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoFuquanDailyReport.Services
{
    public static class FileService
    {
        ///<summary>
        /// 根据文件名的部分获取完整的文件名名称，如果有多个相同的文件，取第一个文件名
        /// </summary>
        /// <param name="folderName">文件夹名称</param>
        /// <param name="fileName">文件名</param>
        /// <param name="fileExtension">后缀名</param>
        /// <returns></returns>
        public static string GetFileName(string folderName, string fileName, string fileExtension)
        {
            var dirs = Directory.GetFiles($@"{folderName}\", $"*{fileName}*.{fileExtension}");    //结果含有路径

            return dirs[0];
        }
    }
}

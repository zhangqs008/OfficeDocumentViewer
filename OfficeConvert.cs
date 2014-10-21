using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using ICSharpCode.SharpZipLib.Zip;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;

namespace Whir.Software.DocumentViewer
{
    /// <summary>
    ///     Office转换辅助类
    /// </summary>
    public class OfficeConverter
    {
        #region LogCategory enum

        public enum LogCategory
        {
        }

        #endregion

        protected static string CommonFileInclude
        {
            get
            {
                return
                    string.Format("<link type='text/css' rel='stylesheet' href='../../style/common.css'/>" +
                                  Environment.NewLine +
                                  "<script type='text/javascript' src='../../scripts/jquery-1.7.1.js'></script>" +
                                  Environment.NewLine +
                                  "<script type='text/javascript' src='../../scripts/common.js'></script>" +
                                  Environment.NewLine
                        );
            }
        }

        /// <summary>
        ///     转换Word文档成Html文档
        /// </summary>
        /// <param name="sourcePath"></param>
        /// <param name="targetPath"></param>
        /// <returns></returns>
        public static ConvertResult WordToHtml(string sourcePath, string targetPath)
        {
            try
            {
                var wordApp = new Application();
                Document currentDoc = wordApp.Documents.Open(sourcePath);
                currentDoc.SaveAs(targetPath, WdSaveFormat.wdFormatFilteredHTML);
                currentDoc.Close();
                wordApp.Quit();
                ChanageCharset(targetPath);
                return new ConvertResult { IsSuccess = true, Message = targetPath };
            }
            catch (Exception ex)
            {
                Log("转换Word文档成Html文档时异常", ex.Message + ex.StackTrace);
                return new ConvertResult { IsSuccess = false, Message = ex.Message + ex.StackTrace };
            }
        }

        /// <summary>
        ///     转换Excel文档成Html
        /// </summary>
        /// <param name="sourcePath"></param>
        /// <param name="targetPath"></param>
        /// <returns></returns>
        public static ConvertResult ExcelToHtml(string sourcePath, string targetPath)
        {
            try
            {
                var excelApp = new Microsoft.Office.Interop.Excel.Application();
                Workbook workbook = excelApp.Application.Workbooks.Open(sourcePath);
                workbook.SaveAs(targetPath, XlFileFormat.xlHtml);
                workbook.Close();
                excelApp.Quit();
                ChanageCharset(targetPath);
                return new ConvertResult { IsSuccess = true, Message = targetPath };
            }
            catch (Exception ex)
            {
                Log("转换Excel文档成Html文档时异常", ex.Message + ex.StackTrace);
                return new ConvertResult { IsSuccess = false, Message = ex.Message + ex.StackTrace };
            }
        }

        /// <summary>
        ///     转换ppt文档成Html文档
        /// </summary>
        /// <param name="sourcePath"></param>
        /// <param name="targetPath"></param>
        /// <returns></returns>
        public static ConvertResult PptToHtml(string sourcePath, string targetPath)
        {
            try
            {
                #region 先转换为图片，再由图片组成Html页面

                var sourceFile = new FileInfo(sourcePath);
                var powerPoint = new Microsoft.Office.Interop.PowerPoint.Application();
                Presentation open = powerPoint.Presentations.Open(sourcePath, MsoTriState.msoTrue,
                                                                  MsoTriState.msoFalse, MsoTriState.msoFalse);

                //注意：有些版本的PowerPoint(如：Office 2013 Professional)不能保存为Html，
                //所以，先保存为图片，再由图片组成一个Html页面来预览
                open.SaveAs(targetPath, PpSaveAsFileType.ppSaveAsJPG, MsoTriState.msoTrue);

                var targetDirPath =
                    new DirectoryInfo(string.Format("{0}\\{1}", Path.GetDirectoryName(targetPath), sourceFile.Name));
                const string template =
                    @"<!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Transitional//EN' 'http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd'>
                        <html xmlns='http://www.w3.org/1999/xhtml'>
                        <meta http-equiv=Content-Type content='text/html; charset=gbk'>
                        {0}
                        <head>
                            <title>{1}</title>
                        </head>
                        <body>
                            <div class='file_content'>
                                {2}
                            </div>
                        </body>
                        </html>";
                var images = new StringBuilder();
                foreach (FileInfo file in targetDirPath.GetFiles())
                {
                    images.AppendFormat("<p class='images'><img class='pptImages' src='{0}/{1}' /></p>{2}",
                                        targetDirPath.Name, file.Name, Environment.NewLine);
                }
                WriteFile(targetPath, string.Format(template, CommonFileInclude, sourceFile, images));
                open.Close();
                powerPoint.Quit();
                return new ConvertResult { IsSuccess = true, Message = targetPath };

                #endregion
            }
            catch (Exception ex)
            {
                Log("转换PPT文档成Html文档时异常", ex.Message + ex.StackTrace);
                return new ConvertResult { IsSuccess = false, Message = ex.Message + ex.StackTrace };
            }
        }

        /// <summary>
        ///     转换图片成Html文档
        /// </summary>
        /// <param name="sourcePath">相对路径</param>
        /// <param name="targetPath"></param>
        /// <returns></returns>
        public static ConvertResult ImageToHtml(string sourcePath, string targetPath)
        {
            try
            {
                var file = new FileInfo(sourcePath);
                const string htmlTemplate =
                    @"<!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Transitional//EN' 'http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd'>
                                           <html xmlns='http://www.w3.org/1999/xhtml'>
                                           <meta http-equiv=Content-Type content='text/html; charset=gbk'>
                                           {0}
                                            <head>
                                                 <title>{1}</title>
                                            </head>
                                            <body>
                                                <div class='file_content'>
                                                    <img class='file_image' src='{1}' />
                                                </div>
                                            </body>
                                          </html>";
                string content = string.Format(htmlTemplate, CommonFileInclude, file.Name);
                WriteFile(targetPath, content);
                return new ConvertResult { IsSuccess = true, Message = targetPath };
            }
            catch (Exception ex)
            {
                Log("转换PPT文档成Html文档时异常", ex.Message + ex.StackTrace);
                return new ConvertResult { IsSuccess = false, Message = ex.Message + ex.StackTrace };
            }
        }

        /// <summary>
        ///     转换压缩包成Html文档
        /// </summary>
        /// <param name="sourcePath">相对路径</param>
        /// <param name="targetPath"></param>
        /// <returns></returns>
        public static ConvertResult ZipToHtml(string sourcePath, string targetPath)
        {
            try
            {
                var htmlFiles = new List<FileInfo>();
                var inputFile = new FileInfo(sourcePath);

                #region 第一步：解压文件

                var files = new List<FileInfo>();
                switch (inputFile.Extension.Replace(".", "").ToLower())
                {
                    case "rar":
                        files = UnRarFile(sourcePath);
                        break;
                    case "zip":
                        files = UnZipFile(sourcePath);
                        break;
                }

                #endregion

                if (files.Count > 0)
                {
                    #region 第二步：转换文件

                    foreach (FileInfo fileInfo in files)
                    {
                        var result = new ConvertResult();
                        string tempPath = fileInfo.FullName + ".htm";
                        switch (fileInfo.Extension.Replace(".", "").ToLower())
                        {
                            #region Word转换

                            case "doc":
                            case "docx":
                            case "txt":
                            case "csv":
                            case "cs":
                            case "wps":
                            case "js":
                            case "xml":
                            case "config":
                                result = WordToHtml(fileInfo.FullName, tempPath);
                                break;

                            #endregion

                            #region Excel转换

                            case "xls":
                            case "xlsx":
                            case "et":
                                result = ExcelToHtml(fileInfo.FullName, tempPath);
                                break;

                            #endregion

                            #region PPT转换

                            case "ppt":
                            case "pptx":
                            case "wpp":
                            case "dps":

                                result = PptToHtml(fileInfo.FullName, tempPath);
                                break;

                            #endregion

                            #region 图片转换

                            case "jpg":
                            case "png":
                            case "ico":
                            case "gif":
                            case "bmp":
                                result = ImageToHtml(fileInfo.FullName, tempPath);
                                break;

                            #endregion
                        }
                        if (result.IsSuccess)
                        {
                            htmlFiles.Add(new FileInfo(result.Message));
                        }
                    }

                    #endregion

                    #region 第三步：组装Html页面

                    const string template =
                        @"<!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Transitional//EN' 'http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd'>
                                           <html xmlns='http://www.w3.org/1999/xhtml'>
                                           <meta http-equiv=Content-Type content='text/html; charset=gbk'> 
                                            {0}
                                            <head>
                                                 <title>{1}</title>
                                            </head>
                                            <body>
                                                <div class='file_content'>
                                                    <ul class='file_rar_files_ul'>
                                                      {2}
                                                   </ul>
                                               </div>
                                            </body>
                                          </html>";
                    string fileHtml = "";
                    foreach (FileInfo htmlFile in htmlFiles)
                    {
                        //去掉转换文档的后缀名，得到解压后的原始文档文件路径
                        var unPressSourceFile = new FileInfo(htmlFile.Name.Replace(htmlFile.Extension, ""));
                        if (htmlFile.Directory != null)
                            fileHtml +=
                                string.Format(
                                    "<li class='file_rar_files_li'><img src='../../style/images/16x16/{0}.gif' style='vertical-align: middle;' />&nbsp;<a href='{1}/{2}'>{3}</a></li>",
                                    unPressSourceFile.Extension.Replace(".", ""),
                                    htmlFile.Directory.Name,
                                    htmlFile.Name,
                                    htmlFile.Name.Replace(htmlFile.Extension, "") + Environment.NewLine);
                    }
                    string content = string.Format(template, CommonFileInclude, inputFile.Name, fileHtml);
                    WriteFile(targetPath, content);
                    return new ConvertResult { IsSuccess = true, Message = targetPath };

                    #endregion
                }
                return new ConvertResult { IsSuccess = false, Message = "压缩包内无任何可预览的文件！" };
            }
            catch (Exception ex)
            {
                Log("转换Rar文档成Html文档时异常", ex.Message + ex.StackTrace);
                return new ConvertResult { IsSuccess = false, Message = ex.Message + ex.StackTrace };
            }
        }


        /// <summary>
        ///     日志记录
        /// </summary>
        /// <param name="title"></param>
        /// <param name="content"></param> 
        public static void Log(string title, string content)
        {
            var logPath = HttpContext.Current.Server.MapPath("~/log.txt");
            using (StreamWriter w = File.AppendText(logPath))
            {
                w.WriteLine("# " + DateTime.Now.ToString("yyyy-MM-dd_HH:mm:ss ") + content);
                w.Close();
            }
        }

        /// <summary>
        ///     写入文件
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="content"></param>
        public static void WriteFile(string filePath, string content)
        {
            try
            {
                var fs = new FileStream(filePath, FileMode.Create);
                Encoding encode = Encoding.GetEncoding("gb2312");
                //获得字节数组
                byte[] data = encode.GetBytes(content);
                //开始写入
                fs.Write(data, 0, data.Length);
                //清空缓冲区、关闭流
                fs.Flush();
                fs.Close();
            }
            catch (Exception ex)
            {
                Log("修改文件编码异常", ex.Message);
            }
        }

        /// <summary>
        ///     读取文件
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static string ReadFile(string filePath)
        {
            string str;
            using (var sr = new StreamReader(filePath, Encoding.Default))
            {
                str = sr.ReadToEnd();
            }
            return str;
        }

        /// <summary>
        ///     修改生成文件charset,防止框架页乱码
        /// </summary>
        /// <param name="filepath"></param>
        private static void ChanageCharset(string filepath)
        {
            string content = ReadFile(filepath);
            var regex = new Regex("(?i)<meta\\s*http-equiv[^>]*?>", RegexOptions.CultureInvariant | RegexOptions.Compiled);
            content = regex.Replace(content, "<meta http-equiv=Content-Type content=\"text/html; charset=gbk\">" +
                                    CommonFileInclude);
            content = content.Replace("position: absolute;", "").Replace("position:absolute;", "");
            WriteFile(filepath, content);
        }

        #region 文件解压

        /// <summary>
        ///     文件解压（Zip格式）
        /// </summary>
        /// <param name="zipFilePath"></param>
        /// <returns></returns>
        private static List<FileInfo> UnZipFile(string zipFilePath)
        {
            var files = new List<FileInfo>();
            var zipFile = new FileInfo(zipFilePath);
            if (!File.Exists(zipFilePath))
            {
                return files;
            }

            using (var zipInputStream = new ZipInputStream(File.OpenRead(zipFilePath)))
            {
                ZipEntry theEntry;
                while ((theEntry = zipInputStream.GetNextEntry()) != null)
                {
                    if (zipFilePath != null)
                    {
                        string dir = Path.GetDirectoryName(zipFilePath);
                        if (dir != null)
                        {
                            string dirName = Path.Combine(dir + "\\ConvertHtml",
                                                          zipFile.Name.Replace(zipFile.Extension, ""));
                            string fileName = Path.GetFileName(theEntry.Name);

                            if (!string.IsNullOrEmpty(dirName))
                            {
                                if (!Directory.Exists(dirName))
                                {
                                    Directory.CreateDirectory(dirName);
                                }
                            }
                            if (!string.IsNullOrEmpty(fileName))
                            {
                                string filePath = Path.Combine(dirName, theEntry.Name);
                                using (FileStream streamWriter = File.Create(filePath))
                                {
                                    var data = new byte[2048];
                                    while (true)
                                    {
                                        int size = zipInputStream.Read(data, 0, data.Length);
                                        if (size > 0)
                                        {
                                            streamWriter.Write(data, 0, size);
                                        }
                                        else
                                        {
                                            break;
                                        }
                                    }
                                }
                                files.Add(new FileInfo(filePath));
                            }
                        }
                    }
                }
            }
            return files;
        }

        /// <summary>
        ///     文件解压（Rar格式）
        /// </summary>
        /// <param name="rarFilePath"></param>
        /// <returns></returns>
        public static List<FileInfo> UnRarFile(string rarFilePath)
        {
            var files = new List<FileInfo>();
            var fileInput = new FileInfo(rarFilePath);
            if (fileInput.Directory != null)
            {
                string dirName = Path.Combine(fileInput.Directory.FullName + "\\ConvertHtml",
                                              fileInput.Name.Replace(fileInput.Extension, ""));
                if (!string.IsNullOrEmpty(dirName))
                {
                    if (!Directory.Exists(dirName))
                    {
                        Directory.CreateDirectory(dirName);
                    }
                }
                dirName = dirName.EndsWith("\\") ? dirName : dirName + "\\"; //最后这个斜杠不能少！
                string shellArguments = string.Format("x -o+ {0} {1}", rarFilePath, dirName);
                using (var unrar = new Process())
                {
                    unrar.StartInfo.FileName = @"C:\Program Files (x86)\WinRAR\WinRAR.exe"; //WinRar安装路径！
                    unrar.StartInfo.Arguments = shellArguments; //隐藏rar本身的窗口
                    unrar.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    unrar.Start();
                    unrar.WaitForExit(); //等待解压完成
                    unrar.Close();
                }
                var dir = new DirectoryInfo(dirName);
                files.AddRange(dir.GetFiles());
            }
            return files;
        }

        #endregion

        #region Rar格式解压

        ///// <summary>
        ///// 文件解壓（rar格式）使用SharpCompress组件 需.net 3.5以上才支持！
        ///// </summary>
        ///// <param name="rarFilePath"></param>
        ///// <returns></returns>
        //private static List<FileInfo> UnRarFile(string rarFilePath)
        //{
        //    var files = new List<FileInfo>();
        //    if (File.Exists(rarFilePath))
        //    {
        //        var fileInput = new FileInfo(rarFilePath);
        //        using (Stream stream = File.OpenRead(rarFilePath))
        //        {
        //            var reader = ReaderFactory.Open(stream);
        //            if (fileInput.Directory != null)
        //            {
        //                string dirName = Path.Combine(fileInput.Directory.FullName, fileInput.Name.Replace(fileInput.Extension, ""));

        //                if (!string.IsNullOrEmpty(dirName))
        //                {
        //                    if (!Directory.Exists(dirName))
        //                    {
        //                        Directory.CreateDirectory(dirName);
        //                    }
        //                }
        //                while (reader.MoveToNextEntry())
        //                {
        //                    if (!reader.Entry.IsDirectory)
        //                    {
        //                        reader.WriteEntryToDirectory(dirName, ExtractOptions.ExtractFullPath | ExtractOptions.Overwrite);
        //                        files.Add(new FileInfo(reader.Entry.FilePath));
        //                    }
        //                }
        //            }
        //        }
        //    }
        //    return files;
        //}

        #endregion

        #region Nested type: ConvertResult

        /// <summary>
        ///     文档转换结果
        /// </summary>
        public class ConvertResult
        {
            public bool IsSuccess { get; set; }
            public string Message { get; set; }
        }

        #endregion
    }
}
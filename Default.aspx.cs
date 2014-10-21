using System;
using System.Globalization;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.UI;

namespace Whir.Software.DocumentViewer
{
    public partial class Default : Page
    {
        protected string DocumentDirName = "Documents";

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                string url = Request.QueryString["url"];
                if (!string.IsNullOrEmpty(url))
                {
                    if (new Regex(@"(?i)/.*\.[a-zA-Z]{3,}").IsMatch(url))
                    {
                        string extension = url.Substring(url.LastIndexOf('.'));
                        string fileName = url.Substring(url.LastIndexOf('/') + 1);
                        string filePath = Path.Combine(Server.MapPath("~/" + DocumentDirName + "/"), fileName);
                        string targetConvertDirPath = Server.MapPath(string.Format("~/{0}/ConvertHtml", DocumentDirName));
                        //目标文件路径
                        var targetConvertFilePath = string.Format("{0}/ConvertHtml/{1}.htm", DocumentDirName, fileName);
                        var targetPath = Server.MapPath("~/" + targetConvertFilePath);
                        if (File.Exists(Server.MapPath("~/" + targetConvertFilePath)))
                        {
                            #region 如果文件已存在

                            Uri uri = HttpContext.Current.Request.Url;
                            string port = uri.Port == 80 ? string.Empty : ":" + uri.Port;
                            string webUrl = string.Format("{0}://{1}{2}/", uri.Scheme, uri.Host, port);
                            //Response.Redirect("Preview.aspx?url=" + webUrl + targetConvertFilePath + "&source=" + url);
                            Response.Redirect(string.Format("Preview.aspx?url={0}{1}&source={2}", webUrl, (BasePath + targetConvertFilePath).Replace("//", "/").TrimStart('/'), url));
                              

                            #endregion
                        }
                        else
                        {
                            #region 第一步：下载文件

                            try
                            {
                                var webClient = new WebClient();
                                if (!Directory.Exists(Server.MapPath("~/" + DocumentDirName + "/")))
                                {
                                    Directory.CreateDirectory(Server.MapPath("~/" + DocumentDirName + "/"));
                                }
                                webClient.DownloadFile(url, filePath);
                            }
                            catch (Exception ex)
                            {
                                ResponseMsg(false, ex.Message.ToString(CultureInfo.InvariantCulture));
                            }
                            #endregion

                            if (File.Exists(filePath))
                            {
                                #region  第二步：转换文件


                                string sourcePath = filePath;
                                if (!Directory.Exists(targetConvertDirPath))
                                {
                                    Directory.CreateDirectory(targetConvertDirPath);
                                }
                                if (File.Exists(targetConvertFilePath))
                                {
                                    File.Delete(targetConvertFilePath);
                                }

                                OfficeConverter.ConvertResult result;
                                switch (extension.Replace(".", "").ToLower())
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
                                        result = OfficeConverter.WordToHtml(sourcePath, targetPath);
                                        break;

                                    #endregion

                                    #region Excel转换

                                    case "xls":
                                    case "xlsx":
                                    case "et":
                                        result = OfficeConverter.ExcelToHtml(sourcePath, targetPath);
                                        break;

                                    #endregion

                                    #region PPT转换

                                    case "ppt":
                                    case "pptx":
                                    case "wpp":
                                    case "dps":

                                        result = OfficeConverter.PptToHtml(sourcePath, targetPath);
                                        break;

                                    #endregion

                                    #region 图片转换

                                    case "jpg":
                                    case "png":
                                    case "ico":
                                    case "gif":
                                    case "bmp":
                                        result = OfficeConverter.ImageToHtml(sourcePath, targetPath);
                                        break;

                                    #endregion

                                    #region 压缩包

                                    case "zip":
                                    case "rar":
                                        result = OfficeConverter.ZipToHtml(sourcePath, targetPath);
                                        break;

                                    #endregion

                                    default:
                                        result = new OfficeConverter.ConvertResult
                                                     {
                                                         IsSuccess = false,
                                                         Message = "该文档类型不支持在线预览！"
                                                     };
                                        break;
                                }
                                if (result.IsSuccess)
                                {
                                    Uri uri = HttpContext.Current.Request.Url;
                                    string port = uri.Port == 80 ? string.Empty : ":" + uri.Port;
                                    string webUrl = string.Format("{0}://{1}{2}/", uri.Scheme, uri.Host, port);
                                    Response.Redirect(string.Format("Preview.aspx?url={0}{1}&source={2}", webUrl, (BasePath + targetConvertFilePath).Replace("//", "/").TrimStart('/'), url));
                                }
                                else
                                {
                                    ResponseMsg(false, "对不起，" + result.Message);
                                }


                                #endregion
                            }
                            else
                            {
                                ResponseMsg(false, "对不起，文件下载失败，未找到对应文件！");
                            }
                        }
                    }
                    else
                    {
                        ResponseMsg(false, "对不起，文件路径不正确！");
                    }
                }
            }
        }

        protected void ResponseMsg(bool isOk, string msg)
        {
            Response.Clear();
            string msgage = "<span style='font-size:12px;color:" + (isOk ? "green" : "red") + "'>" + msg + "</span>";
            Response.Write(msgage);
            Response.End();
        }
        public static string BasePath
        {
            get { return VirtualPathUtility.AppendTrailingSlash(HttpContext.Current.Request.ApplicationPath); }
        }
    }
}
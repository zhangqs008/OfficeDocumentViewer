<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="Whir.Software.DocumentViewer.Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <title>Office文档在线预览</title>
    <link href="style/common.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">
    <div class="guide">
        <h3>
            Document Viewer-Office文档在线预览</h3>
    </div>
    <div class="main">
        使用说明：请在网站地址后，用url参数传入您想预览的文件路径。<br />
        如：/Default.aspx<span style="   color: red">?url=http://192.168.0.253:8010/Task/Editor/eWebeditor/uploadfile/20140120094053568.doc</span>
    </div>
    </form>
</body>
</html>

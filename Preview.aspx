<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Preview.aspx.cs" Inherits="Whir.Software.DocumentViewer.Preview" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>文档预览</title>
    <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
    <link href="style/common.css" rel="stylesheet" type="text/css" />
    <script src="Scripts/jquery-1.7.1.js" type="text/javascript"></script>
</head>
<body>
    <form id="form1" runat="server">
    <div class="guide" style="width: 90%">
        <a href="Default.aspx">返回首页</a> &nbsp;<a href="<%=Source %>">下载文件</a>
    </div>
    <div class="main" style="width: 90%">
        <iframe src="<%=Url %>" style="height: 500px; padding: 10px; width: 98%;" id="ifmEditPage"
            marginheight="0" marginwidth="0" frameborder="0" height="0px" width="0px"></iframe>
    </div>
    </form>
	
    <script type="text/javascript">
        function onload() {
            var winHeight = $(window).height();
            var winWidth = $(window).width(); 
            var width = winWidth;
            var height = winHeight - 90;

            //$("#ifmEditPage").width(width);
            $("#ifmEditPage").height(height);
        }
        window.onresize = onload;
        $(function () {
            onload();
        });
    </script>

</body>
</html>

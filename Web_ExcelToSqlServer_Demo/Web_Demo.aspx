<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Web_Demo.aspx.cs" Inherits="Web_ExcelToSqlServer_Demo.Web_Demo" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:FileUpload ID="FileUpload1" runat="server" />
        <br />
        <asp:Button ID="Button1" runat="server" Text="上传" onclick="Button1_Click" />
    </div>
    </form>
</body>
</html>

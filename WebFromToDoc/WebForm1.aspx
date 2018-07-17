<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="WebForm1.aspx.cs" Inherits="WebFromToDoc.WebForm1" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <div>
            Name
            <asp:TextBox runat="server" ID="txtName" required />
            <br />
            Title
            <asp:TextBox runat="server" ID="txtTitle" required />
            <br />
            <br />
            Excel File
            <asp:DropDownList ID="DropDownList1" AutoPostBack="true" runat="server" OnSelectedIndexChanged="DropDownList1_SelectedIndexChanged"></asp:DropDownList>
            <br />
            <asp:GridView ID="GridView1" runat="server">
            </asp:GridView>
            <br />
            <asp:Button runat="server" Text="Ok" OnClick="Unnamed_Click" />
            <br />
            <br />
        </div>
    </form>
</body>
</html>

<%@ Page Language="C#" AutoEventWireup="true" CodeFile="CS.aspx.cs" Inherits="_Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title></title>
    <style type="text/css">
        body
        {
            font-family: Arial;
            font-size: 10pt;
        }
        td, th
        {
            height: 25px;
            width: 100px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
    <div>
<%--    <asp:FileUpload ID="FileUpload1" runat="server" />
        <asp:Button ID="btnUpload" runat="server" Text="Upload" OnClick="Upload" />
        <hr />--%>
        <asp:Label ID="lblmsg" runat="server"></asp:Label><br />
        <asp:GridView ID="GridView1" runat="server" HeaderStyle-BackColor="#3AC0F2" HeaderStyle-ForeColor="White"
            RowStyle-BackColor="#A1DCF2" AlternatingRowStyle-BackColor="White" AlternatingRowStyle-ForeColor="#000"
            OnRowDeleting="GridView1_RowDeleting" DataKeyNames="Id" AutoGenerateColumns="false">
            <Columns>
                <asp:TemplateField HeaderText="Delete">
                <ItemTemplate>
                    <asp:ImageButton ID="ImageButton1" runat="server" ImageUrl="~/img/Remove.gif" CommandName="Delete"
                        OnClientClick="return confirm('Are you sure want to delete record?')" />
                </ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="Name" HeaderText="File Name" />
                <asp:TemplateField ItemStyle-HorizontalAlign="Center">
                    <ItemTemplate>
                        <asp:LinkButton ID="lnkDownload" runat="server" Text="Download" OnClick="DownloadFile"
                            CommandArgument='<%# Eval("Id") %>'></asp:LinkButton>
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
        </asp:GridView>
    </div>
    </form>
</body>
</html>

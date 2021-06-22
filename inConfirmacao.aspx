<%@ Page Language="VB" AutoEventWireup="false" CodeFile="inConfirmacao.aspx.vb" Inherits="inConfirmacao" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Confirmar</title>
</head>
<body style="font-size: 8pt; color: navy; font-family: Verdana, Arial; left: 68px; width: 80px; position: relative; height: 376px;" leftmargin="0" rightmargin="0" topmargin="20" bottommargin="15" bgcolor="#cfd7e5">
    <form id="form1" runat="server">
    <div align="center">
        &nbsp;<img src="Images/Tit_Negociacao2.jpg" /><br />
        <br />
        <br />
        <table id="TABLE1" border="1" cellpadding="1" cellspacing="0" style="width: 376px; height: 32px;" bordercolor="#ffffff">
            <tr>
                <td align="center" bgcolor="gainsboro" colspan="3" rowspan="1" style="padding-right: 10px;
                    padding-left: 10px; font-weight: bold; font-size: 10pt; clip: rect(auto auto auto auto);
                    color: navy; height: 30px">
                    Confirmação de Negociação</td>
            </tr>
            <tr>
                <td align="center" bgcolor="azure" colspan="2" rowspan="2" style="width: 243px; color: white;
                    height: 30px; background-color: #5e75a1">
                    <b>Acordo</b></td>
                <td align="left" bgcolor="gainsboro" colspan="1" rowspan="2" style="padding-right: 10px;
                    padding-left: 10px; width: 201px; clip: rect(auto auto auto auto); height: 30px">
                    <asp:Label ID="lblParcela" runat="server"></asp:Label></td>
            </tr>
            <tr>
            </tr>
            <tr>
                <td align="center" colspan="3" rowspan="1" style="padding-right: 10px; padding-left: 10px;
                    clip: rect(auto auto auto auto); height: 30px">
                    <br />
        <asp:Button ID="btnVoltar" runat="server" BackColor="White" BorderColor="DarkGray"
            Font-Names="verdana,arial" Font-Size="8pt" ForeColor="Navy" Text="Voltar" Font-Bold="True" Width="56px" />
                    <asp:Button ID="btnConfirmar" runat="server" BackColor="White" BorderColor="DarkGray"
            Font-Names="verdana,arial" Font-Size="8pt" ForeColor="Navy" Text="Confirmar Negociação" Font-Bold="True" Width="176px" />&nbsp;
                    <Button ID="btnImprimir" runat="server" onclick="javaScript:window.print()" style="border-right: darkgray thin outset; border-top: darkgray thin outset; font-weight: bold; font-size: 8pt; border-left: darkgray thin outset; color: navy; border-bottom: darkgray thin outset; font-family: verdana, Arial; background-color: white" visible="false">
                        Imprimir</Button><br />
                    <br />
                    &nbsp;<asp:LinkButton ID="LinkNew" runat="server">Nova Negociação</asp:LinkButton></td>
            </tr>
        </table>
        </div>
    </form>
</body>
</html>

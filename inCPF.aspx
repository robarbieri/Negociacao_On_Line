<%@ Page Language="VB" AutoEventWireup="true" CodeFile="inCPF.aspx.vb" Inherits="inCPF" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Validação</title>
</head>
<body style="font-size: 8pt; color: navy; font-family: Verdana, Arial" leftmargin="0" rightmargin="0" topmargin="20" bottommargin="15" bgcolor="#cfd7e5">
    <form id="form1" runat="server">
    <div align="center">
        <br />
        <br />
        <img src="Images/TitCPF2.jpg" /><br />
        <br />
        <table align="center" border="1" cellpadding="2" cellspacing="0" style="width: 320px" bordercolor="white">
            <tr>
                <td align="center" colspan="1" rowspan="3" style="width: 236px; color: white; height: 27px;" bgcolor="#5e75a1">
                    CPF<asp:Label ID="lblEMail" runat="server" ForeColor="DarkOrange" Text="*" Width="8px"></asp:Label></td>
                <td align="left" colspan="4" rowspan="3" style="width: 222px; height: 27px;" bgcolor="gainsboro">
                    <asp:TextBox ID="txtCPF" runat="server" Font-Names="verdana,arial" Font-Size="8pt"
                        ForeColor="Navy" MaxLength="14" Width="120px"></asp:TextBox></td>
            </tr>
            <tr>
            </tr>
            <tr>
            </tr>
            <tr>
                <td align="center" bgcolor="#5e75a1" colspan="1" rowspan="1" style="width: 236px; color: white; height: 27px;">
                    Data de Nascimento<asp:Label ID="Label1" runat="server" ForeColor="DarkOrange" Text="*"
                        Width="8px"></asp:Label></td>
                <td align="left" bgcolor="gainsboro" colspan="4" rowspan="1" style="width: 222px; height: 27px;">
                    <asp:TextBox ID="txtDataNasc" runat="server" onKeyDown="javascript:FormatarCampoData();" Font-Names="verdana,arial" Font-Size="8pt"
                        ForeColor="Navy" MaxLength="10" Width="72px"></asp:TextBox>(dd/mm/aaaa)</td>
            </tr>
            <tr>
                <td align="center" bgcolor="#5e75a1" colspan="1" rowspan="1" style="width: 236px;
                    color: white; height: 27px">
                    Código</td>
                <td align="left" bgcolor="gainsboro" colspan="4" rowspan="1" style="width: 222px;
                    height: 27px">
                    <asp:TextBox ID="txtCodigo" runat="server" Font-Names="verdana,arial" Font-Size="8pt"
                        ForeColor="Navy" MaxLength="10" Width="72px"></asp:TextBox></td>
            </tr>
        </table>
        <br />
        &nbsp;<asp:Button ID="btnCPF" runat="server" Font-Names="Verdana,Arial" Font-Size="8pt"
                        ForeColor="Navy" Text="Continuar" BackColor="White" BorderColor="DarkGray" /><br />
        <br />
        <a href="javascript:popupcenter('Carta.htm','POPup',434,547,0,0);">Clique aqui para visualizar a Carta de Autorização</a></div>
    </form>
</body>
</html>
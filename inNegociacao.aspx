<%@ Page Language="VB" AutoEventWireup="true" CodeFile="inNegociacao.aspx.vb" Inherits="inNegociacao" %>

<%@ Register Assembly="System.Web.Extensions, Version=1.0.61025.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"
    Namespace="System.Web.UI" TagPrefix="asp" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Negociação</title>
</head>
<body style="font-size: 8pt; color: navy; font-family: Verdana, Arial; left: 68px; width: 80px; position: relative; height: 376px;" leftmargin="0" rightmargin="0" topmargin="20" bottommargin="15" bgcolor="#cfd7e5">
    <form id="form1" runat="server">
    <asp:ScriptManager ID="ScriptManager" runat="server">
        </asp:ScriptManager>
    <div align="center">
        <asp:UpdatePanel ID="UpdatePanel" runat="server">
            <ContentTemplate>
        <img src="Images/Tit_Negociacao2.jpg" id="IMG1" /><br />
                <div style="width: 280px; height: 1px" align="center">
                    <asp:Label ID="lblMsg" runat="server" BackColor="White" BorderColor="White" BorderStyle="Solid"
                        BorderWidth="10px" ForeColor="Red" Visible="False" Width="356px" Font-Bold="True" Font-Strikeout="False"></asp:Label></div>
        <table id="TABLE1" border="1" cellpadding="1" cellspacing="0" style="width: 376px; height: 56px;" bordercolor="#ffffff">
            <tr>
                <td colspan="2" rowspan="3" style="width: 243px; height: 24px; color: white; background-color: #5e75a1;" align="left" bgcolor="azure">
                    Valor da Dívida</td>
                <td rowspan="3" style="width: 201px; height: 24px; padding-left: 10px;" align="left" bgcolor="gainsboro">
                    <asp:Label ID="lblSaldo" runat="server"></asp:Label></td>
            </tr>
            <tr>
            </tr>
            <tr>
            </tr>
            <tr>
                <td align="left" colspan="2" rowspan="1" style="width: 243px; height: 24px; color: white; background-color: #5e75a1;" bgcolor="azure">
                    <strong>
                    Condição<br />
                    </strong>
                    (Parc / Valor Desconto)</td>
                <td align="left" rowspan="1" style="width: 201px; height: 33px; padding-left: 10px;" bgcolor="gainsboro">
                    <asp:DropDownList ID="cboCondicao" runat="server" Width="208px" Font-Names="verdana,arial" Font-Size="8pt" ForeColor="Navy">
                    </asp:DropDownList></td>
            </tr>
            <tr>
                <td align="left" bgcolor="azure" colspan="2" rowspan="1" style="width: 243px; height: 24px; color: white; background-color: #5e75a1;">
                    <strong>
                    Vencimento 1ª Parcela<br />
                    </strong>
                    (DD/MM/AAAA)</td>
                <td align="left" bgcolor="gainsboro" rowspan="1" style="width: 201px; height: 33px; padding-left: 10px;">
                    <asp:TextBox ID="txtVencto" runat="server" Width="72px" onKeyDown="javascript:FormatarCampoData();" MaxLength="10" Font-Names="verdana,arial" Font-Size="8pt" ForeColor="Navy"></asp:TextBox>&nbsp;
                </td>
            </tr>
            <tr>
                <td align="left" bgcolor="azure" colspan="2" rowspan="1" style="width: 243px; height: 24px; color: white; background-color: #5e75a1;">
                    Entrada <b>Mínima</b></td>
                <td align="left" bgcolor="gainsboro" rowspan="1" style="width: 201px; height: 24px; padding-left: 10px;">
                    <asp:TextBox ID="txtValor" runat="server" Width="72px" onKeyDown="javascript:FormataValor('txtValor',10,event)" onKeyUp="javascript:ReFormataValor('txtValor',10,event)" Font-Names="verdana,arial" Font-Size="8pt" ForeColor="Navy" MaxLength="15" Enabled="False"></asp:TextBox></td>
            </tr>
            <tr>
                <td align="center" bgcolor="azure" colspan="2" rowspan="1" style="width: 243px; height: 24px; color: white; background-color: #5e75a1;">
                    <b>Acordo</b></td>
                <td align="left" bgcolor="gainsboro" rowspan="1" style="width: 201px; height: 24px; padding-right: 10px; padding-left: 10px; clip: rect(auto auto auto auto);" colspan="">
                    <asp:UpdateProgress ID="UpdateProgress" runat="server" AssociatedUpdatePanelID="UpdatePanel"
                        DisplayAfter="1">
                        <ProgressTemplate>
                            <div id="Progress" style="font-size: 7pt; color: Navy; font-family: Verdana,Arial" align="center">
                                <img src="Images/Loading1.gif" />&nbsp; Calculando...
                            </div>
                        </ProgressTemplate>
                    </asp:UpdateProgress>
                    <asp:Label ID="lblParcela" runat="server"></asp:Label></td>
            </tr>
            <tr>
                <td align="center" bgcolor="#cfd7e5" colspan="3" rowspan="1" style="padding-right: 10px;
                    padding-left: 10px; clip: rect(auto auto auto auto); height: 47px">
                    &nbsp;<br />
                    <%--<button ID="btnCalcular" onclick="javascript:verificaForm()" style="font-size: 8pt;font-family: Verdana,Arial; height: 20px; border-left-color: gray; border-bottom-color: gray; color: navy; border-top-style: solid; border-top-color: gray; border-right-style: solid; border-left-style: solid; border-right-color: gray; border-bottom-style: solid; background-color: white;" type="button">Calcular</button>--%>
                    <asp:Button ID="btnCalcularServer" runat="server" BackColor="White" BorderColor="DarkGray"
            Font-Names="verdana,arial" Font-Size="8pt" ForeColor="Navy" Text="Calcular" Font-Bold="False" Width="176px" /><br />
                    <br />
        <asp:Button ID="btnNegociacao" runat="server" BackColor="White" BorderColor="DarkGray"
            Font-Names="verdana,arial" Font-Size="8pt" ForeColor="Navy" Text="Registrar Negociação" Font-Bold="True" Width="176px" /></td>
            </tr>
        </table>
            </ContentTemplate>
        </asp:UpdatePanel>
        &nbsp;&nbsp;&nbsp;</div>
<%--        <asp:TextBox ID="Atualizado" runat="server" Height="1px" Width="1px" BackColor="Transparent" BorderColor="Transparent" Enabled="False" ForeColor="Transparent"></asp:TextBox>
        <asp:TextBox ID="Principal" runat="server" Height="1px" Width="1px" BackColor="Transparent" BorderColor="Transparent" Enabled="False" ForeColor="Transparent"></asp:TextBox>
        <asp:TextBox ID="PercentHonor" runat="server" Height="1px" Width="1px" BackColor="Transparent" BorderColor="Transparent" Enabled="False" ForeColor="Transparent"></asp:TextBox>
        <asp:TextBox ID="txtidLogin" runat="server" BackColor="Transparent" BorderColor="Transparent"
            Enabled="False" ForeColor="Transparent" Height="1px" Width="1px"></asp:TextBox>
        <asp:TextBox ID="txtContratoOS" runat="server" BackColor="Transparent" BorderColor="Transparent"
            Enabled="False" ForeColor="Transparent" Height="1px" Width="1px"></asp:TextBox>--%>
    </form>
</body>
</html>

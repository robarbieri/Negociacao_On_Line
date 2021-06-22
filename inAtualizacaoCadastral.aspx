<%@ Page Language="VB" AutoEventWireup="false" CodeFile="inAtualizacaoCadastral.aspx.vb" Inherits="inAtualizacaoCadastral" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Untitled Page</title>
</head>
<body style="font-size: 8pt; color: navy; font-family: Verdana, Arial; left: 63px; width: 400px; position: relative; height: 376px;" leftmargin="0" rightmargin="0" topmargin="20" bottommargin="15" bgcolor="#cfd7e5">
    <form id="form1" runat="server">
    <div align="center">
        <img src="Images/TitAtzCadastral2.jpg" /><br />
        <table id="TABLE1" border="1" cellpadding="1" cellspacing="0" style="width: 392px; height: 376px;" bordercolor="#ffffff">
            <tr>
                <td colspan="2" rowspan="3" style="width: 172px; color: white; background-color: #5e75a1;" align="left" bgcolor="azure">
                    E-Mail<asp:Label ID="lblEMail" runat="server" ForeColor="DarkOrange" Text="*"
                        Width="8px"></asp:Label></td>
                <td rowspan="3" style="width: 194px; height: 16px;" align="left" bgcolor="gainsboro">
                    &nbsp;<asp:TextBox ID="txtEMail" runat="server" Height="16px" MaxLength="50" Width="176px"></asp:TextBox></td>
            </tr>
            <tr>
            </tr>
            <tr>
            </tr>
            <tr>
                <td align="left" colspan="2" rowspan="1" style="width: 172px; color: white; background-color: #5e75a1;" bgcolor="azure">
                    Endereço <b>Residencial</b>
                    <asp:Label ID="lblEnd1" runat="server" ForeColor="DarkOrange" Text="*"
                        Width="8px"></asp:Label></td>
                <td align="left" rowspan="1" style="width: 194px; height: 22px" bgcolor="gainsboro">
                    &nbsp;<asp:TextBox ID="txtEndResLogr" runat="server" Height="16px" MaxLength="50" Width="176px"></asp:TextBox></td>
            </tr>
            <tr>
                <td align="left" colspan="2" rowspan="1" style="width: 172px; color: white; background-color: #5e75a1;" bgcolor="azure">
                    Número
                    <asp:Label ID="lblEnd2" runat="server" ForeColor="DarkOrange" Text="*"
                        Width="8px"></asp:Label></td>
                <td align="left" rowspan="1" style="width: 194px; height: 22px" bgcolor="gainsboro">
                    &nbsp;<asp:TextBox ID="txtEndResNum" runat="server" Height="16px" MaxLength="6" Width="56px"></asp:TextBox></td>
            </tr>
            <tr>
                <td align="left" colspan="2" rowspan="1" style="width: 172px; color: white; background-color: #5e75a1;" bgcolor="azure">
                    Complemento
                    </td>
                <td align="left" rowspan="1" style="width: 194px; height: 22px" bgcolor="gainsboro">
                    &nbsp;<asp:TextBox ID="txtEndResCompl" runat="server" Height="16px" MaxLength="20" Width="80px"></asp:TextBox></td>
            </tr>
            <tr>
                <td align="left" colspan="2" rowspan="1" style="width: 172px; color: white; background-color: #5e75a1;" bgcolor="azure">
                    Bairro
                    <asp:Label ID="lblEnd3" runat="server" ForeColor="DarkOrange" Text="*"
                        Width="8px"></asp:Label></td>
                <td align="left" rowspan="1" style="width: 194px; height: 22px" bgcolor="gainsboro">
                    &nbsp;<asp:TextBox ID="txtEndResBairro" runat="server" Height="16px" MaxLength="50" Width="176px"></asp:TextBox></td>
            </tr>
            <tr>
                <td align="left" colspan="2" rowspan="1" style="width: 172px; color: white; background-color: #5e75a1;" bgcolor="azure">
                    Cep
                    <asp:Label ID="lblEnd4" runat="server" ForeColor="DarkOrange" Text="*"
                        Width="8px"></asp:Label></td>
                <td align="left" rowspan="1" style="width: 194px; height: 22px" bgcolor="gainsboro">
                    &nbsp;<asp:TextBox ID="txtEndResCEP" runat="server" Height="16px" MaxLength="9" Width="80px"></asp:TextBox></td>
            </tr>
            <tr>
                <td align="left" colspan="2" rowspan="1" style="width: 172px; color: white; background-color: #5e75a1;" bgcolor="azure">
                    Cidade
                    <asp:Label ID="lblEnd5" runat="server" ForeColor="DarkOrange" Text="*"
                        Width="8px"></asp:Label></td>
                <td align="left" rowspan="1" style="width: 194px; height: 22px" bgcolor="gainsboro">
                    &nbsp;<asp:TextBox ID="txtEndResCidade" runat="server" Height="16px" MaxLength="40" Width="160px"></asp:TextBox></td>
            </tr>
            <tr>
                <td align="left" colspan="2" rowspan="1" style="width: 172px; color: white; background-color: #5e75a1;" bgcolor="azure">
                    UF
                    <asp:Label ID="lblEnd6" runat="server" ForeColor="DarkOrange" Text="*"
                        Width="8px"></asp:Label></td>
                <td align="left" rowspan="1" style="width: 194px; height: 22px" bgcolor="gainsboro">
                    &nbsp;<asp:DropDownList ID="cboResUF" runat="server" Font-Names="verdana,arial" Font-Size="8pt"
                        ForeColor="Navy" Width="48px">
                    </asp:DropDownList></td>
            </tr>
            <tr>
                <td align="left" bgcolor="azure" colspan="2" rowspan="2" style="width: 172px; color: white;
                    background-color: #5e75a1">
                    Telefone <b>Residencial</b><asp:Label ID="lblFone" runat="server" ForeColor="DarkOrange" Text="*"
                        Width="8px"></asp:Label></td>
                <td align="left" bgcolor="gainsboro" rowspan="1" style="width: 194px; height: 25px">
                    &nbsp;DDD
                    <asp:TextBox ID="txtDDDRes1" runat="server" Height="16px" MaxLength="2" Width="32px"></asp:TextBox>&nbsp;
                    Fone
                    <asp:TextBox ID="txtFoneRes1" runat="server" Height="16px" MaxLength="9" Width="72px"></asp:TextBox></td>
            </tr>
            <tr>
                <td align="left" bgcolor="gainsboro" rowspan="1" style="width: 194px; height: 25px">
                    &nbsp;DDD
                    <asp:TextBox ID="txtDDDRes2" runat="server" Height="16px" MaxLength="2" Width="32px"></asp:TextBox>&nbsp;&nbsp;Fax&nbsp;
                    &nbsp;<asp:TextBox ID="txtFoneRes2" runat="server" Height="16px" MaxLength="9" Width="72px"></asp:TextBox></td>
            </tr>
            <tr>
                <td align="left" colspan="2" rowspan="1" style="width: 172px; color: white; background-color: #5e75a1;" bgcolor="azure">
                    Endereço <b>Comercial</b></td>
                <td align="left" rowspan="1" style="width: 194px; height: 25px" bgcolor="gainsboro">
                    &nbsp;<asp:TextBox ID="txtEndComLogr" runat="server" Height="16px" MaxLength="50" Width="176px"></asp:TextBox></td>
            </tr>
            <tr>
                <td align="left" colspan="2" rowspan="1" style="width: 172px; color: white; background-color: #5e75a1;" bgcolor="azure">
                    Número</td>
                <td align="left" rowspan="1" style="width: 194px; height: 25px" bgcolor="gainsboro">
                    &nbsp;<asp:TextBox ID="txtEndComNum" runat="server" Height="16px" MaxLength="6" Width="32px"></asp:TextBox></td>
            </tr>
            <tr>
                <td align="left" colspan="2" rowspan="1" style="width: 172px; color: white; background-color: #5e75a1;" bgcolor="azure">
                    Complemento</td>
                <td align="left" rowspan="1" style="width: 194px; height: 25px" bgcolor="gainsboro">
                    &nbsp;<asp:TextBox ID="txtEndComCompl" runat="server" Height="16px" MaxLength="20" Width="80px"></asp:TextBox></td>
            </tr>
            <tr>
                <td align="left" colspan="2" rowspan="1" style="width: 172px; color: white; background-color: #5e75a1;" bgcolor="azure">
                    Bairro</td>
                <td align="left" rowspan="1" style="width: 194px; height: 25px" bgcolor="gainsboro">
                    &nbsp;<asp:TextBox ID="txtEndComBairro" runat="server" Height="16px" MaxLength="50" Width="176px"></asp:TextBox></td>
            </tr>
            <tr>
                <td align="left" colspan="2" rowspan="1" style="width: 172px; color: white; background-color: #5e75a1;" bgcolor="azure">
                    Cep</td>
                <td align="left" rowspan="1" style="width: 194px; height: 25px" bgcolor="gainsboro">
                    &nbsp;<asp:TextBox ID="txtEndComCep" runat="server" Height="16px" MaxLength="9" Width="80px"></asp:TextBox></td>
            </tr>
            <tr>
                <td align="left" colspan="2" rowspan="1" style="width: 172px; color: white; background-color: #5e75a1;" bgcolor="azure">
                    Cidade</td>
                <td align="left" rowspan="1" style="width: 194px; height: 25px" bgcolor="gainsboro">
                    &nbsp;<asp:TextBox ID="txtEndComCidade" runat="server" Height="16px" MaxLength="40" Width="160px"></asp:TextBox></td>
            </tr>
            <tr>
                <td align="left" colspan="2" rowspan="1" style="width: 172px; color: white; background-color: #5e75a1;" bgcolor="azure">
                    UF</td>
                <td align="left" rowspan="1" style="width: 194px; height: 25px" bgcolor="gainsboro">
                    &nbsp;<asp:DropDownList ID="cboComUF" runat="server" Font-Names="verdana,arial" Font-Size="8pt"
                        ForeColor="Navy" Width="48px">
                    </asp:DropDownList></td>
            </tr>
            <tr>
                <td align="left" bgcolor="azure" colspan="2" rowspan="2" style="width: 172px; color: white;
                    background-color: #5e75a1">
                    Telefone <b>Comercial</b></td>
                <td align="left" bgcolor="gainsboro" rowspan="1" style="width: 194px; height: 25px">
                    &nbsp;DDD
                    <asp:TextBox ID="txtDDDCom1" runat="server" Height="16px" MaxLength="2" Width="32px"></asp:TextBox>&nbsp;
                    Fone
                    <asp:TextBox ID="txtFoneCom1" runat="server" Height="16px" MaxLength="9" Width="72px"></asp:TextBox></td>
            </tr>
            <tr>
                <td align="left" bgcolor="gainsboro" rowspan="1" style="width: 194px; height: 25px">
                    &nbsp;DDD
                    <asp:TextBox ID="txtDDDCom2" runat="server" Height="16px" MaxLength="2" Width="32px"></asp:TextBox>&nbsp;&nbsp;Fax&nbsp;
                    &nbsp;<asp:TextBox ID="txtFoneCom2" runat="server" Height="16px" MaxLength="9" Width="72px"></asp:TextBox></td>
            </tr>
            <tr>
                <td colspan="2" rowspan="2" style="width: 172px; color: white; background-color: #5e75a1;" align="left" bgcolor="azure">
                    Telefone <b>Celular</b></td>
                <td rowspan="1" style="width: 194px; height: 25px" align="left" bgcolor="gainsboro">
                    &nbsp;DDD
                    <asp:TextBox ID="txtDDDCel1" runat="server" Height="16px" MaxLength="2" Width="32px"></asp:TextBox>&nbsp;
                    Fone
                    <asp:TextBox ID="txtFoneCel1" runat="server" Height="16px" MaxLength="9" Width="72px"></asp:TextBox></td>
            </tr>
            <tr>
                <td rowspan="1" style="width: 194px; height: 25px" align="left" bgcolor="gainsboro">
                    &nbsp;DDD
                    <asp:TextBox ID="txtDDDCel2" runat="server" Height="16px" MaxLength="2" Width="32px"></asp:TextBox>&nbsp;
                    Fone
                    <asp:TextBox ID="txtFoneCel2" runat="server" Height="16px" MaxLength="9" Width="72px"></asp:TextBox></td>
            </tr>
        </table>
        <br />
        <asp:Button ID="btnEndereco" runat="server" BackColor="White" BorderColor="DarkGray"
            Font-Names="verdana,arial" Font-Size="8pt" ForeColor="Navy" Text="Continuar" /></div>
    </form>
</body>
</html>

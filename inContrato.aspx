<%@ Page Language="VB" AutoEventWireup="false" CodeFile="inContrato.aspx.vb" Inherits="inContrato" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Untitled Page</title>
</head>
<body style="font-size: 8pt; color: navy; font-family: Verdana, Arial" leftmargin="0" rightmargin="0" topmargin="20" bottommargin="15" bgcolor="#cfd7e5">
    <form id="form1" runat="server">
    <div align="center">
        <br />
        <br />
        <img src="Images/TitContrato2.jpg" /><br />
        <br />
        <table align="center" border="1" cellpadding="2" cellspacing="0" style="width: 352px" bordercolor="white">
            <tr>
                <td align="center" colspan="1" rowspan="3" style="width: 87px; color: white; background-color: #5e75a1;" bgcolor="#f0ffff">
                    Cartão:</td>
                <td align="center" colspan="4" rowspan="3" style="width: 36px" bgcolor="gainsboro">
                    <asp:DropDownList ID="cboContratos" runat="server" Font-Names="verdana,arial" Font-Size="8pt"
                        ForeColor="Navy" Width="160px">
                    </asp:DropDownList></td>
                <td align="center" colspan="1" rowspan="3" style="width: 35px">
                    <asp:Button ID="btnContrato" runat="server" Font-Names="Verdana,Arial" Font-Size="8pt"
                        ForeColor="Navy" Text="Continuar" BackColor="White" BorderColor="DarkGray" /></td>
            </tr>
            <tr>
            </tr>
            <tr>
            </tr>
        </table>
    
    </div>
    </form>
</body>
</html>

<%

'****************************************************************
'                           Neo
'****************************************************************
' Módulo: Boleto Bancário Lazer - boleto_lazer.asp
' Autor : Adriano Rocha Lima e Silva Araujo
' Inicio: 27/12/2002	11:00
' Fim   : 27/12/2002	00:00
'
' Atualização: 
'       Autor: 
'       Data : 
'
'****************************************************************

Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","no-store"
Response.CacheControl = "no-cache"
Response.Expires = -100000

If Session("Func_ID") = "" Then
	Response.Redirect "../reset.asp"
	Response.End	
End If


Dim BD					'Conexão ao banco de dados. 
Dim rsBanco
Dim rsEscritorio
Dim rsParcelas
Dim vParcelas
Dim vContratante
Dim rsChequeInfo
vContratante = Request.Form("CONT_ID")
if vContratante = "" then
	vContratante = Request.QueryString("CONT_ID")
	if vContratante = "" then
		vContratante = 0
	end if
end if


vTipoBoleto = Request.Form("TipoBoleto")

Set BD = Server.CreateObject("ADODB.Connection")
BD.Open Application("Conexao")

Set rsEscritorio = BD.Execute("SELECT * FROM Escritorio WITH (NOLOCK)")

Dim rsTemp2
Dim vContr 
dim vEndFilial

%>
<html>
<head>
<STYLE type=text/css>
.ti { FONT: 9px Arial, Helvetica, sans-serif }
.ct { FONT: 9px Arial Narrow; COLOR: navy }
.cn { FONT: 9px Arial; COLOR: black }
.cn2 { FONT: 3px Arial; COLOR: black }
.cn3 { FONT: 6px Arial; COLOR: black }
.cp { FONT: bold 11px Arial; COLOR: black }
.ld { FONT: bold 15px Arial; COLOR: #000000 }
.bc { FONT: bold 18px Arial; COLOR: #000000 }
.pb { PAGE-BREAK-AFTER: always }
.texto1 {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 8pt; }
.cn5 { FONT: 11px Arial; COLOR: black }
.texto2
{
	padding-right: 0px;
	padding-left: 0px;
	font-size: 7pt;
	padding-bottom: 0px;
	margin: 0px;
	border-top-style: none;
	padding-top: 0px;
	font-family: Arial, Helvetica, sans-serif;
	border-right-style: none;
	border-left-style: none;
	border-bottom-style: none;
}
.tabela {  background-color: #FFFFFF; border: thin #333333 solid}
</STYLE>
<title>Impressão de Boletos</title>
</head>
<!-- #include file="..\boleto_funcoes.asp" -->
<!-- #include file="..\funcoes.asp" -->
<!-- #include file="..\log_operacoes.asp" -->
<!-- #include file="desc_boleto_laser.asp" -->
<body>

<div id="aguarde" style="position:absolute; width:100%; left: 0px; top: 0px; overflow: auto;">
<br>
<br>
<br>
<br>
	<table border=0 cellspacing=0 cellpadding=0 align=center ID="Table17">
	  <tr>
		 <td valign=top class=texto1><img src="../images/status_wait.gif"></td>
	  </tr>
	</table>
</div>
<CENTER>
<%
Response.Flush 

Dim rsEnderecoDevedor
Dim atab(99)
Dim vQtdBoletos
Dim rsTemp
Dim vValorTaxaBoleto
Dim vCART_ID, vCTRA_ID
Dim vMensagemJuros

vMensagemJuros = false
if Request.Form("rdMensagemJuros") = "" then
	vMensagemJuros = true
end if
vValorTaxaBoleto = Request.Form("ValorTaxaBoleto")
if Request.Form("rdCobraBoleto") = "Não" then
	vValorTaxaBoleto = 0
end if

vQtdboletos = 0
Dim vE

vE = Request.Form("cboEndereco")
vCTRA_ID = Request.Form("ctra_id")

Dim vFilialDestino, rsFilial, rsFilial2
Dim vTel, vFax
Dim rsCPF, vCPF, vValidadeBoleto, rsNumAco
Dim vInicioSeqBoleto, vTitulos

Dim vBoletos, vCONT_ID, vAcordoNumero, vTpBoleto
vBoletos = ""

if Request.QueryString("recuperador") <> "" then
	vTpBoleto = "A"
	if Request.QueryString("individual") <> "" then
		Set rsParcelas = BD.Execute("SELECT * FROM Boletos_Recuperador b WITH (NOLOCK) WHERE BORE_ID = " & Request.QueryString("individual"))
	else
		Set rsParcelas = BD.Execute("SELECT * FROM Boletos_Recuperador_Temp bt WITH (NOLOCK) JOIN Boletos_Recuperador b WITH (NOLOCK) ON bt.BORE_ID = b.BORE_ID WHERE bt.FUNC_ID = " & Session("FUNC_ID"))
	end if
	Do While Not rsParcelas.EOF
		Set rsFilial = BD.Execute("SELECT ca.CONT_ID, FILI_ID FROM Carteiras ca WITH (NOLOCK) JOIN Contratos c WITH (NOLOCK) ON c.CART_ID = ca.CART_ID WHERE CTRA_ID = " & rsParcelas("CTRA_ID"))
		vFilialDestino = rsFilial("FILI_ID")

		vCONT_ID = rsFilial("CONT_ID")
		
		rsFilial.Close
		Set rsFilial = Nothing
		'Laser Acordo
		
		Set rsCPF = BD.Execute("SELECT DEVE_CGCCPF, CTRA_Numero FROM Devedores d WITH (NOLOCK) JOIN Contratos c WITH (NOLOCK) ON c.DEVE_ID = d.DEVE_ID WHERE CTRA_ID = " & rsParcelas("CTRA_ID"))
		if vCONT_ID = 74 or vCONT_ID = 75 then
			if len(rsCPF("DEVE_CGCCPF")) = 14 then
				vCPF = "0" & Mid(rsCPF("DEVE_CGCCPF"), 1, 8)
			else
				vCPF = Mid(rsCPF("DEVE_CGCCPF"), 1, 9)
			end if
		elseif vCONT_ID = 69 then
			vCPF = CDbl(rsCPF("DEVE_CGCCPF"))
			vCPF = Mid(vCPF & "0000", 1, 4)
		elseif vCONT_ID = 77 then
			vCPF = rsCPF("CTRA_Numero")
			vCPF = Right("00000000000000" & vCPF, 14)
		elseif vCONT_ID = 83 then
			vCPF = rsCPF("CTRA_Numero")
		elseif vCONT_ID = 97 then
			vCPF = CDbl(rsCPF("DEVE_CGCCPF"))
			vCPF = "1200905" & Mid(vCPF & "000", 1, 3)
		else
			vCPF = cdbl(rsCPF("DEVE_CGCCPF"))
			vCPF = Mid(vCPF & "00000", 1, 5)
		end if
		
		Set rsBanco = BD.Execute("SELECT CONT_InicioSeqBoleto FROM Contratante WITH (NOLOCK) WHERE CONT_ID = " & vCONT_ID)
		if not isnull(rsBanco("CONT_InicioSeqBoleto")) then
			vInicioSeqBoleto = rsBanco("CONT_InicioSeqBoleto") 'Substituir por campo no banco de dados
		else
			vInicioSeqBoleto = 0
		end if

		if vCONT_ID = 11 then
			Set rsNumAco = BD.Execute("SELECT CTRA_NumeroAcordoContratante FROM Contratos c WITH (NOLOCK) WHERE CTRA_ID = " & rsParcelas("CTRA_ID"))
			if isnull(rsNumAco("CTRA_NumeroAcordoContratante")) then
				Set rsNumAco = BD.Execute("SELECT MAX(CTRA_NumeroAcordoContratante) + 1 NumAcordo FROM Contratos c WITH (NOLOCK) JOIN Carteiras ca WITH (NOLOCK) ON c.CART_ID = ca.CART_ID WHERE CONT_ID = " & vCONT_ID)
				BD.Execute "UPDATE Contratos SET CTRA_NumeroAcordoContratante = " & rsNumAco("NumAcordo") & " WHERE CTRA_ID = " & rsParcelas("CTRA_ID") 
				PreencheLOG "UPDATE Contratos SET CTRA_NumeroAcordoContratante = " & rsNumAco("NumAcordo") & " WHERE CTRA_ID = " & rsParcelas("CTRA_ID") , vFilialDestino
			end if
			Set rsNumAco = BD.Execute("SELECT CTRA_NumeroAcordoContratante FROM Contratos c WITH (NOLOCK) WHERE CTRA_ID = " & rsParcelas("CTRA_ID"))
            if not isnull(rsNumAco("CTRA_NumeroAcordoContratante")) then
				vAcordoNumero = Right("000000" & rsNumAco("CTRA_NumeroAcordoContratante"), 6)
			end if
			vNumDoc = vAcordoNumero & "001" & Right("00" & rsParcelas("BORE_QtdParcelas"), 2)
		elseif vCONT_ID = 14 then
			Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT BOGE_Numero NumDocumento FROM Boletos_Gerados b WITH (NOLOCK) JOIN Contratos c WITH (NOLOCK) ON b.CTRA_ID = c.CTRA_ID JOIN Carteiras ca WITH (NOLOCK) ON c.CART_ID = ca.CART_ID WHERE CONT_ID = " & vCONT_ID & " AND BOGE_Numero >= " & vInicioSeqBoleto & " AND Len(BOGE_Numero) = 9) res")
			if IsNull(rsTemp("NumDocumento")) then
				vNumDoc = vInicioSeqBoleto
			else
				vNumDoc = rsTemp("NumDocumento")
			end if
		elseif vCONT_ID = 74 or vCONT_ID = 75 then
			Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT PARC_NumDocumento NumDocumento FROM Parcelas WITH (NOLOCK) WHERE PARC_NumDocumento LIKE '" & CDbl(vCPF) & "%' AND FPAG_ID = 3 AND Len(PARC_NumDocumento) = Len(" & CDbl(vCPF) & ") + 4 UNION SELECT BOAV_NumDocumento NumDocumento FROM Boletos_Avulsos WITH (NOLOCK) WHERE BOAV_NumDocumento LIKE '" & CDbl(vCPF) & "%' AND Len(BOAV_NumDocumento) = Len(" & CDbl(vCPF) & ") + 4 UNION SELECT BOGE_Numero NumDocumento FROM Boletos_Gerados WITH (NOLOCK) WHERE BOGE_Numero LIKE '" & CDbl(vCPF) & "%' AND Len(BOGE_Numero) = Len(" & CDbl(vCPF) & ") + 4) res")
			if IsNull(rsTemp("NumDocumento")) then
				vNumDoc = vCPF & "0001"
			else
				vNumDoc = rsTemp("NumDocumento")
			end if
		elseif vCONT_ID = 69 then
			Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT PARC_NumDocumento NumDocumento FROM Parcelas WITH (NOLOCK) WHERE PARC_NumDocumento LIKE '" & vCPF & "%' AND FPAG_ID = 3 AND Len(PARC_NumDocumento) = 7 UNION SELECT BOAV_NumDocumento NumDocumento FROM Boletos_Avulsos WITH (NOLOCK) WHERE BOAV_NumDocumento LIKE '" & vCPF & "%' AND Len(BOAV_NumDocumento) = 7 UNION SELECT BOGE_Numero NumDocumento FROM Boletos_Gerados WITH (NOLOCK) WHERE BOGE_Numero LIKE '" & vCPF & "%' AND Len(BOGE_Numero) = 7) res")
			if IsNull(rsTemp("NumDocumento")) then
				vNumDoc = vCPF & "001"
			else
				vNumDoc = rsTemp("NumDocumento")
			end if
		elseif vCONT_ID = 77 then
			vNumDoc = vCPF
		elseif vCONT_ID = 83 then
			if Month(rsParcelas("BORE_Vencimento")) = 10 then
				vNumDoc = "0"
			elseif Month(rsParcelas("BORE_Vencimento")) = 11 then
				vNumDoc = "5"
			elseif Month(rsParcelas("BORE_Vencimento")) = 12 then
				vNumDoc = "6"
			else
				vNumDoc = CSTR(Month(rsParcelas("BORE_Vencimento")))
			end if
			'Falta identificar o produto
			Set rsTemp = BD.Execute("SELECT BAND_Numero FROM Bandeiras b WITH (NOLOCK) JOIN Contratos c WITH (NOLOCK) ON b.BAND_ID = c.BAND_ID WHERE CTRA_ID = " & rsParcelas("CTRA_ID"))
			if rsTemp.EOF then
				vNumDoc = vNumDoc & "001" & Mid(vCPF, 7, 7)
			else
				vNumDoc = vNumDoc & Right("000" & rsTemp("BAND_Numero"), 3) & Mid(vCPF, 7, 7)
			end if
		elseif vCONT_ID = 97 then
			Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT PARC_NumDocumento NumDocumento FROM Parcelas WITH (NOLOCK) WHERE PARC_NumDocumento LIKE '" & vCPF & "%' AND FPAG_ID = 3 AND Len(PARC_NumDocumento) = 13 UNION SELECT BOAV_NumDocumento NumDocumento FROM Boletos_Avulsos WITH (NOLOCK) WHERE BOAV_NumDocumento LIKE '" & vCPF & "%' AND Len(BOAV_NumDocumento) = 13 UNION SELECT BOGE_Numero NumDocumento FROM Boletos_Gerados WITH (NOLOCK) WHERE BOGE_Numero LIKE '" & vCPF & "%' AND Len(BOGE_Numero) = 13) res")
			if IsNull(rsTemp("NumDocumento")) then
				vNumDoc = vCPF & "001"
			else
				vNumDoc = rsTemp("NumDocumento")
			end if
		else
			Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT PARC_NumDocumento NumDocumento FROM Parcelas WITH (NOLOCK) WHERE PARC_NumDocumento LIKE '" & vCPF & "%' AND FPAG_ID = 3 AND Len(PARC_NumDocumento) = 8 UNION SELECT BOAV_NumDocumento NumDocumento FROM Boletos_Avulsos WITH (NOLOCK) WHERE BOAV_NumDocumento LIKE '" & vCPF & "%' AND Len(BOAV_NumDocumento) = 8 UNION SELECT BOGE_Numero NumDocumento FROM Boletos_Gerados WITH (NOLOCK) WHERE BOGE_Numero LIKE '" & vCPF & "%' AND Len(BOGE_Numero) = 8) res")
			'Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT PARC_NumDocumento NumDocumento FROM Parcelas WITH (NOLOCK) WHERE PARC_NumDocumento LIKE '" & vCPF & "%' AND FPAG_ID = 3 AND Len(PARC_NumDocumento) = 8 UNION SELECT BOAV_NumDocumento NumDocumento FROM Boletos_Avulsos WITH (NOLOCK) WHERE BOAV_NumDocumento LIKE '" & vCPF & "%' AND Len(BOAV_NumDocumento) = 8) res")
			'Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT PARC_NumDocumento NumDocumento FROM Parcelas WITH (NOLOCK) WHERE PARC_NumDocumento LIKE '" & Session("FILI_ID") & "%' AND FPAG_ID = 3 AND Len(PARC_NumDocumento) >= 8 AND Len(PARC_NumDocumento) <=9 UNION SELECT BOAV_NumDocumento NumDocumento FROM Boletos_Avulsos WITH (NOLOCK) WHERE BOAV_NumDocumento LIKE '" & Session("FILI_ID") & "%' AND Len(BOAV_NumDocumento) >= 8 AND Len(BOAV_NumDocumento) <=9) res")
			'Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT PARC_NumDocumento NumDocumento FROM Parcelas WITH (NOLOCK) WHERE PARC_NumDocumento LIKE '" & Session("FILI_ID") & "%' AND FPAG_ID = 3 AND Len(PARC_NumDocumento) >= 7 UNION SELECT BOAV_NumDocumento NumDocumento FROM Boletos_Avulsos WITH (NOLOCK) WHERE BOAV_NumDocumento LIKE '" & Session("FILI_ID") & "%' AND Len(BOAV_NumDocumento) >= 7) res")
			if IsNull(rsTemp("NumDocumento")) then
				vNumDoc = vCPF & "001"
			else
				vNumDoc = rsTemp("NumDocumento")
			end if
		end if
		
		if IsNull(rsParcelas("BORE_Titulos")) or rsParcelas("BORE_Titulos") = "" then
			vTitulos = "NULL"
		else
			vTitulos = "'" & rsParcelas("BORE_Titulos") & "'"
		end if

		vID = RetornaProximoNumero("Boletos_Avulsos", "BOAV_ID", Session("FILI_ID"))
		if Request.QueryString("individual") <> "" then
			BD.Execute "INSERT INTO Boletos_Avulsos Values(" & vID & ", " & rsParcelas("ctra_id") & ", " & rsParcelas("edev_id") & ", '" & vNumDoc & "', 'A', " & Replace(Replace(rsParcelas("BORE_Valor"),".",""),",",".") & ", '" & rsParcelas("BORE_Vencimento") & "', " & rsParcelas("BORE_QtdParcelas") & ", '" & rsParcelas("BORE_Vencimento") & "', 0, NULL, " & vTitulos & ", NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL)"
			PreencheLOG "INSERT INTO Boletos_Avulsos Values(" & vID & ", " & rsParcelas("ctra_id") & ", " & rsParcelas("edev_id") & ", '" & vNumDoc & "', 'A', " & Replace(Replace(rsParcelas("BORE_Valor"),".",""),",",".") & ", '" & rsParcelas("BORE_Vencimento") & "', " & rsParcelas("BORE_QtdParcelas") & ", '" & rsParcelas("BORE_Vencimento") & "', 0, NULL, " & vTitulos & ", NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL)", vFilialDestino
		else
			BD.Execute "INSERT INTO Boletos_Avulsos Values(" & vID & ", " & rsParcelas("ctra_id") & ", " & rsParcelas("edev_id") & ", '" & vNumDoc & "', 'A', " & Replace(Replace(rsParcelas("BORE_Valor"),".",""),",",".") & ", '" & rsParcelas("BORE_Vencimento") & "', " & rsParcelas("BORE_QtdParcelas") & ", '" & rsParcelas("BORE_Validade") & "', 0, NULL, " & vTitulos & ", NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL)"
			PreencheLOG "INSERT INTO Boletos_Avulsos Values(" & vID & ", " & rsParcelas("ctra_id") & ", " & rsParcelas("edev_id") & ", '" & vNumDoc & "', 'A', " & Replace(Replace(rsParcelas("BORE_Valor"),".",""),",",".") & ", '" & rsParcelas("BORE_Vencimento") & "', " & rsParcelas("BORE_QtdParcelas") & ", '" & rsParcelas("BORE_Validade") & "', 0, NULL, " & vTitulos & ", NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL)", vFilialDestino
		end if

		BD.Execute "INSERT INTO Boletos_Gerados(CTRA_ID, BOGE_Numero, BOGE_TipoBoleto) VALUES(" & rsParcelas("CTRA_ID") & ", " & vNumDoc & ", 'A')"
		PreencheLOG "INSERT INTO Boletos_Gerados(CTRA_ID, BOGE_Numero, BOGE_TipoBoleto) VALUES(" & rsParcelas("CTRA_ID") & ", " & vNumDoc & ", 'A')", vFilialDestino

		BD.Execute "UPDATE Boletos_Recuperador SET BORE_Enviado = 1 WHERE BORE_ID = " & rsParcelas("BORE_ID")
		PreencheLOG "UPDATE Boletos_Recuperador SET BORE_Enviado = 1 WHERE BORE_ID = " & rsParcelas("BORE_ID"), vFilialDestino

		vBoletos = vBoletos & vID & ","

		rsParcelas.MoveNext
	Loop
	vBoletos = left(vBoletos, len(vBoletos) - 1)

	Set rsParcelas = BD.Execute("SELECT 0 PARC_ID, b.CTRA_ID, d.DEVE_ID, DEVE_Nome, DEVE_CGCCPF, BOAV_NumDocumento PARC_NumDocumento, EDEV_ID, CTRA_Numero, b.BOAV_QtdParcelas Plano, BOAV_TaxaBoleto, BOAV_Validade, BOAV_Validade, BOAV_Vencimento, BOAV_Valor, c.CART_ID, CTRA_VencDebito, SCON_ID, CTRA_Conta, CART_Descricao, BOAV_Texto2 FROM Boletos_Avulsos b WITH (NOLOCK) JOIN Contratos c WITH (NOLOCK) ON b.CTRA_ID = c.CTRA_ID JOIN Devedores d WITH (NOLOCK) ON c.DEVE_ID = d.DEVE_ID JOIN Carteiras ca WITH (NOLOCK) ON c.CART_ID = ca.CART_ID WHERE BOAV_ID IN (" & vBoletos & ") AND ACMD_ID IS NULL ORDER BY CART_Descricao, DEVE_Nome ")
elseif Request.QueryString("automatico") <> "" then
	vTpBoleto = "P"
	Set rsParcelas = BD.Execute("SELECT p.ACOR_ID, PARC_ID, PARC_NumDocumento, FPAG_ID, CTRA_ID, PARC_Numero, Qtd Plano, PARC_Vencimento FROM Parcelas p WITH (NOLOCK) JOIN Acordos a WITH (NOLOCK) ON a.ACOR_ID = p.ACOR_ID JOIN (SELECT ACOR_ID, COUNT(*) Qtd FROM Parcelas WITH (NOLOCK) GROUP BY ACOR_ID) r ON r.ACOR_ID = p.ACOR_ID WHERE PARC_ID IN (SELECT PARC_ID FROM Boletos_Parcela_Temp WITH (NOLOCK) WHERE FUNC_ID = " & Session("FUNC_ID") & ")")
	Do While Not rsParcelas.EOF
		Set rsFilial = BD.Execute("SELECT ca.CONT_ID, FILI_ID FROM Carteiras ca WITH (NOLOCK) JOIN Contratos c WITH (NOLOCK) ON c.CART_ID = ca.CART_ID WHERE CTRA_ID = " & rsParcelas("CTRA_ID"))
		vFilialDestino = rsFilial("FILI_ID")
		
		vCONT_ID = rsFilial("CONT_ID")
		
		
		rsFilial.Close
		Set rsFilial = Nothing
		'Laser Acordo
		
		Set rsCPF = BD.Execute("SELECT DEVE_CGCCPF, CTRA_Numero FROM Devedores d WITH (NOLOCK) JOIN Contratos c WITH (NOLOCK) ON c.DEVE_ID = d.DEVE_ID WHERE CTRA_ID = " & rsParcelas("CTRA_ID"))
		if vCONT_ID = 74 or vCONT_ID = 75 then
			if len(rsCPF("DEVE_CGCCPF")) = 14 then
				vCPF = "0" & Mid(rsCPF("DEVE_CGCCPF"), 1, 8)
			else
				vCPF = Mid(rsCPF("DEVE_CGCCPF"), 1, 9)
			end if
		elseif vCONT_ID = 69 then
			vCPF = CDbl(rsCPF("DEVE_CGCCPF"))
			vCPF = Mid(vCPF & "0000", 1, 4)
		elseif vCONT_ID = 77 then
			vCPF = rsCPF("CTRA_Numero")
			vCPF = Right("00000000000000" & vCPF, 14)
		elseif vCONT_ID = 83 then
			vCPF = rsCPF("CTRA_Numero")
		elseif vCONT_ID = 97 then
			vCPF = CDbl(rsCPF("DEVE_CGCCPF"))
			vCPF = "1200905" & Mid(vCPF & "000", 1, 3)
		else
			vCPF = CDbl(rsCPF("DEVE_CGCCPF"))
			vCPF = Mid(vCPF & "00000", 1, 5)
		end if

		Set rsBanco = BD.Execute("SELECT CONT_InicioSeqBoleto FROM Contratante WITH (NOLOCK) WHERE CONT_ID = " & vCONT_ID)
		if not isnull(rsBanco("CONT_InicioSeqBoleto")) then
			vInicioSeqBoleto = rsBanco("CONT_InicioSeqBoleto") 'Substituir por campo no banco de dados
		else
			vInicioSeqBoleto = 0
		end if

		'if (vCONT_ID = 11 and Len(rsParcelas("PARC_NumDocumento")) < 9) or ((rsParcelas("FPAG_ID") = 3 and (IsNull(rsParcelas("PARC_NumDocumento")) or rsParcelas("PARC_NumDocumento") = "" or rsParcelas("PARC_NumDocumento") = "0")) or (rsParcelas("FPAG_ID") <> 3) or IsNull(rsParcelas("FPAG_ID"))) then
		if (vCONT_ID = 11) or ((rsParcelas("FPAG_ID") = 3 and (IsNull(rsParcelas("PARC_NumDocumento")) or rsParcelas("PARC_NumDocumento") = "" or rsParcelas("PARC_NumDocumento") = "0")) or (rsParcelas("FPAG_ID") <> 3) or IsNull(rsParcelas("FPAG_ID"))) then
			if vCONT_ID = 11 then
				Set rsNumAco = BD.Execute("SELECT CTRA_NumeroAcordoContratante FROM Contratos c WITH (NOLOCK) WHERE CTRA_ID = " & rsParcelas("CTRA_ID"))
                if not isnull(rsNumAco("CTRA_NumeroAcordoContratante")) then
					vAcordoNumero = Right("000000" & rsNumAco("CTRA_NumeroAcordoContratante"), 6)
				else
					If Len(rsParcelas("ACOR_ID")) > 5 Then
						vAcordoNumero = Right("000000" & Mid(rsParcelas("ACOR_ID"), 1, Len(rsParcelas("ACOR_ID")) - 3), 6)
					Else
						vAcordoNumero = Right("000000" & rsParcelas("ACOR_ID"), 6)
					End If
				end if
				vNumDoc = vAcordoNumero & "0" & Right("00" & rsParcelas("PARC_Numero"), 2) & Right("00" & rsParcelas("Plano"), 2)
			elseif vCONT_ID = 14 then
				Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT BOGE_Numero NumDocumento FROM Boletos_Gerados b WITH (NOLOCK) JOIN Contratos c WITH (NOLOCK) ON b.CTRA_ID = c.CTRA_ID JOIN Carteiras ca WITH (NOLOCK) ON c.CART_ID = ca.CART_ID WHERE CONT_ID = " & vCONT_ID & " AND BOGE_Numero >= " & vInicioSeqBoleto & " AND Len(BOGE_Numero) = 9) res")
				if IsNull(rsTemp("NumDocumento")) then
					vNumDoc = vInicioSeqBoleto
				else
					vNumDoc = rsTemp("NumDocumento")
				end if
			elseif vCONT_ID = 74 or vCONT_ID = 75 then
				Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT PARC_NumDocumento NumDocumento FROM Parcelas WITH (NOLOCK) WHERE PARC_NumDocumento LIKE '" & CDbl(vCPF) & "%' AND FPAG_ID = 3 AND Len(PARC_NumDocumento) = Len(" & CDbl(vCPF) & ") + 4 UNION SELECT BOAV_NumDocumento NumDocumento FROM Boletos_Avulsos WITH (NOLOCK) WHERE BOAV_NumDocumento LIKE '" & CDbl(vCPF) & "%' AND Len(BOAV_NumDocumento) = Len(" & CDbl(vCPF) & ") + 4 UNION SELECT BOGE_Numero NumDocumento FROM Boletos_Gerados WITH (NOLOCK) WHERE BOGE_Numero LIKE '" & CDbl(vCPF) & "%' AND Len(BOGE_Numero) = Len(" & CDbl(vCPF) & ") + 4) res")
				if IsNull(rsTemp("NumDocumento")) then
					vNumDoc = vCPF & "0001"
				else
					vNumDoc = rsTemp("NumDocumento")
				end if
			elseif vCONT_ID = 69 then
				Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT PARC_NumDocumento NumDocumento FROM Parcelas WITH (NOLOCK) WHERE PARC_NumDocumento LIKE '" & vCPF & "%' AND FPAG_ID = 3 AND Len(PARC_NumDocumento) = 7 UNION SELECT BOAV_NumDocumento NumDocumento FROM Boletos_Avulsos WITH (NOLOCK) WHERE BOAV_NumDocumento LIKE '" & vCPF & "%' AND Len(BOAV_NumDocumento) = 7 UNION SELECT BOGE_Numero NumDocumento FROM Boletos_Gerados WITH (NOLOCK) WHERE BOGE_Numero LIKE '" & vCPF & "%' AND Len(BOGE_Numero) = 7) res")
				if IsNull(rsTemp("NumDocumento")) then
					vNumDoc = vCPF & "001"
				else
					vNumDoc = rsTemp("NumDocumento")
				end if
			elseif vCONT_ID = 77 then
				vNumDoc = vCPF
			elseif vCONT_ID = 83 then
				if Month(rsParcelas("PARC_Vencimento")) = 10 then
					vNumDoc = "0"
				elseif Month(rsParcelas("PARC_Vencimento")) = 11 then
					vNumDoc = "5"
				elseif Month(rsParcelas("PARC_Vencimento")) = 12 then
					vNumDoc = "6"
				else
					vNumDoc = CSTR(Month(rsParcelas("PARC_Vencimento")))
				end if
				Set rsTemp = BD.Execute("SELECT BAND_Numero FROM Bandeiras b WITH (NOLOCK) JOIN Contratos c WITH (NOLOCK) ON b.BAND_ID = c.BAND_ID WHERE CTRA_ID = " & rsParcelas("CTRA_ID"))
				if rsTemp.EOF then
					vNumDoc = vNumDoc & "001" & Mid(vCPF, 7, 7)
				else
					vNumDoc = vNumDoc & Right("000" & rsTemp("BAND_Numero"), 3) & Mid(vCPF, 7, 7)
				end if
			elseif vCONT_ID = 97 then
				Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT PARC_NumDocumento NumDocumento FROM Parcelas WITH (NOLOCK) WHERE PARC_NumDocumento LIKE '" & vCPF & "%' AND FPAG_ID = 3 AND Len(PARC_NumDocumento) = 13 UNION SELECT BOAV_NumDocumento NumDocumento FROM Boletos_Avulsos WITH (NOLOCK) WHERE BOAV_NumDocumento LIKE '" & vCPF & "%' AND Len(BOAV_NumDocumento) = 13 UNION SELECT BOGE_Numero NumDocumento FROM Boletos_Gerados WITH (NOLOCK) WHERE BOGE_Numero LIKE '" & vCPF & "%' AND Len(BOGE_Numero) = 13) res")
				if IsNull(rsTemp("NumDocumento")) then
					vNumDoc = vCPF & "001"
				else
					vNumDoc = rsTemp("NumDocumento")
				end if
			else
				Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT PARC_NumDocumento NumDocumento FROM Parcelas WITH (NOLOCK) WHERE PARC_NumDocumento LIKE '" & vCPF & "%' AND FPAG_ID = 3 AND Len(PARC_NumDocumento) = 8 UNION SELECT BOAV_NumDocumento NumDocumento FROM Boletos_Avulsos WITH (NOLOCK) WHERE BOAV_NumDocumento LIKE '" & vCPF & "%' AND Len(BOAV_NumDocumento) = 8 UNION SELECT BOGE_Numero NumDocumento FROM Boletos_Gerados WITH (NOLOCK) WHERE BOGE_Numero LIKE '" & vCPF & "%' AND Len(BOGE_Numero) = 8) res")
				if IsNull(rsTemp("NumDocumento")) then
					vNumDoc = vCPF & "001"
				else
					vNumDoc = rsTemp("NumDocumento")
				end if
			end if
			BD.Execute "UPDATE Parcelas SET PARC_NumDocumento = " & vNumDoc & ", FPAG_ID = 3 WHERE PARC_ID = " & rsParcelas("PARC_ID")
			PreencheLOG "UPDATE Parcelas SET PARC_NumDocumento = " & vNumDoc & ", FPAG_ID = 3 WHERE PARC_ID = " & rsParcelas("PARC_ID"), vFilialDestino

			BD.Execute "INSERT INTO Boletos_Gerados(CTRA_ID, BOGE_Numero, PARC_ID, BOGE_TipoBoleto) VALUES(" & rsParcelas("CTRA_ID") & ", " & vNumDoc & ", " & rsParcelas("PARC_ID") & ", 'P')"
			PreencheLOG "INSERT INTO Boletos_Gerados(CTRA_ID, BOGE_Numero, PARC_ID, BOGE_TipoBoleto) VALUES(" & rsParcelas("CTRA_ID") & ", " & vNumDoc & ", " & rsParcelas("PARC_ID") & ", 'P')", vFilialDestino
		end if
		BD.Execute "UPDATE Parcelas SET PARC_BoletoEnviado = 1 WHERE PARC_ID = " & rsParcelas("PARC_ID")
		PreencheLOG "UPDATE Parcelas SET PARC_BoletoEnviado = 1 WHERE PARC_ID = " & rsParcelas("PARC_ID"), vFilialDestino
		rsParcelas.MoveNext
	Loop
	Set rsParcelas = BD.Execute("SELECT ACOR_ValorAcordo, CTRA_DataRecebimentoContrato, ACOR_Data, p.PARC_ID, p.ACOR_ID, PARC_Numero, res.Plano, res.ValorAcordo, Convert(char(10), PARC_Vencimento, 103) PARC_Vencimento, PARC_ValorTotal, PARC_NumDocumento, c.CTRA_ID, CTRA_Numero, DEVE_CGCCPF, c.DEVE_ID, DEVE_Nome, EDEV_ID, 0 BOAV_Validade, CART_ID, FNEG_ID, cb.*, SCON_ID, CTRA_Conta FROM Parcelas p WITH (NOLOCK) JOIN Acordos a WITH(NOLOCK) ON p.ACOR_ID = a.ACOR_ID JOIN Contratos c WITH (NOLOCK) ON a.CTRA_ID = c.CTRA_ID JOIN Devedores d WITH (NOLOCK) ON c.DEVE_ID = d.DEVE_ID JOIN (SELECT ACOR_ID, COUNT(*) Plano, SUM(PARC_ValorTotal) ValorAcordo FROM Parcelas WITH (NOLOCK) GROUP BY ACOR_ID) res ON p.ACOR_ID = res.ACOR_ID JOIN Boletos_Parcela_Temp bp WITH (NOLOCK) ON bp.PARC_ID = p.PARC_ID JOIN Configuracoes_de_Boleto cb WITH (NOLOCK) ON bp.COBO_ID = cb.COBO_ID WHERE bp.FUNC_ID = " & Session("FUNC_ID"))
elseif Request.QueryString("atraso") <> "" then
	vTpBoleto = "P"
	Set rsParcelas = BD.Execute("SELECT p.ACOR_ID, PARC_ID, PARC_NumDocumento, FPAG_ID, CTRA_ID, PARC_Numero, Qtd Plano, PARC_Vencimento FROM Parcelas p WITH (NOLOCK) JOIN Acordos a WITH (NOLOCK) ON a.ACOR_ID = p.ACOR_ID JOIN (SELECT ACOR_ID, COUNT(*) Qtd FROM Parcelas WITH (NOLOCK) GROUP BY ACOR_ID) r ON r.ACOR_ID = p.ACOR_ID WHERE PARC_ID IN (SELECT PARC_ID FROM Boletos_Atraso_Temp WITH (NOLOCK) WHERE FUNC_ID = " & Session("FUNC_ID") & ")")
	Do While Not rsParcelas.EOF
		Set rsFilial = BD.Execute("SELECT ca.CONT_ID, FILI_ID FROM Carteiras ca WITH (NOLOCK) JOIN Contratos c WITH (NOLOCK) ON c.CART_ID = ca.CART_ID WHERE CTRA_ID = " & rsParcelas("CTRA_ID"))
		vFilialDestino = rsFilial("FILI_ID")
		
		vCONT_ID = rsFilial("CONT_ID")
		
		rsFilial.Close
		Set rsFilial = Nothing
		'Laser Acordo
		
		Set rsCPF = BD.Execute("SELECT DEVE_CGCCPF, CTRA_Numero FROM Devedores d WITH (NOLOCK) JOIN Contratos c WITH (NOLOCK) ON c.DEVE_ID = d.DEVE_ID WHERE CTRA_ID = " & rsParcelas("CTRA_ID"))
		if vCONT_ID = 74 or vCONT_ID = 75 then
			if len(rsCPF("DEVE_CGCCPF")) = 14 then
				vCPF = "0" & Mid(rsCPF("DEVE_CGCCPF"), 1, 8)
			else
				vCPF = Mid(rsCPF("DEVE_CGCCPF"), 1, 9)
			end if
		elseif vCONT_ID = 69 then
			vCPF = CDbl(rsCPF("DEVE_CGCCPF"))
			vCPF = Mid(vCPF & "0000", 1, 4)
		elseif vCONT_ID = 77 then
			vCPF = rsCPF("CTRA_Numero")
			vCPF = Right("00000000000000" & vCPF, 14)
		elseif vCONT_ID = 83 then
			vCPF = rsCPF("CTRA_Numero")
		elseif vCONT_ID = 97 then
			vCPF = CDbl(rsCPF("DEVE_CGCCPF"))
			vCPF = "1200905" & Mid(vCPF & "000", 1, 3)
		else
			vCPF = CDbl(rsCPF("DEVE_CGCCPF"))
			vCPF = Mid(vCPF & "00000", 1, 5)
		end if

		Set rsBanco = BD.Execute("SELECT CONT_InicioSeqBoleto FROM Contratante WITH (NOLOCK) WHERE CONT_ID = " & vCONT_ID)
		if not isnull(rsBanco("CONT_InicioSeqBoleto")) then
			vInicioSeqBoleto = rsBanco("CONT_InicioSeqBoleto") 'Substituir por campo no banco de dados
		else
			vInicioSeqBoleto = 0
		end if

		'if (vCONT_ID = 11 and Len(rsParcelas("PARC_NumDocumento")) < 9) or ((rsParcelas("FPAG_ID") = 3 and (IsNull(rsParcelas("PARC_NumDocumento")) or rsParcelas("PARC_NumDocumento") = "" or rsParcelas("PARC_NumDocumento") = "0")) or (rsParcelas("FPAG_ID") <> 3) or IsNull(rsParcelas("FPAG_ID"))) then
		if (vCONT_ID = 11) or ((rsParcelas("FPAG_ID") = 3 and (IsNull(rsParcelas("PARC_NumDocumento")) or rsParcelas("PARC_NumDocumento") = "" or rsParcelas("PARC_NumDocumento") = "0")) or (rsParcelas("FPAG_ID") <> 3) or IsNull(rsParcelas("FPAG_ID"))) then
			if vCONT_ID = 11 then
				Set rsNumAco = BD.Execute("SELECT CTRA_NumeroAcordoContratante FROM Contratos c WITH (NOLOCK) WHERE CTRA_ID = " & rsParcelas("CTRA_ID"))
                if not isnull(rsNumAco("CTRA_NumeroAcordoContratante")) then
					vAcordoNumero = Right("000000" & rsNumAco("CTRA_NumeroAcordoContratante"), 6)
				else
					If Len(rsParcelas("ACOR_ID")) > 5 Then
						vAcordoNumero = Right("000000" & Mid(rsParcelas("ACOR_ID"), 1, Len(rsParcelas("ACOR_ID")) - 3), 6)
					Else
						vAcordoNumero = Right("000000" & rsParcelas("ACOR_ID"), 6)
					End If
				end if
				vNumDoc = vAcordoNumero & "0" & Right("00" & rsParcelas("PARC_Numero"), 2) & Right("00" & rsParcelas("Plano"), 2)
			elseif vCONT_ID = 14 then
				Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT BOGE_Numero NumDocumento FROM Boletos_Gerados b WITH (NOLOCK) JOIN Contratos c WITH (NOLOCK) ON b.CTRA_ID = c.CTRA_ID JOIN Carteiras ca WITH (NOLOCK) ON c.CART_ID = ca.CART_ID WHERE CONT_ID = " & vCONT_ID & " AND BOGE_Numero >= " & vInicioSeqBoleto & " AND Len(BOGE_Numero) = 9) res")
				if IsNull(rsTemp("NumDocumento")) then
					vNumDoc = vInicioSeqBoleto
				else
					vNumDoc = rsTemp("NumDocumento")
				end if
			elseif vCONT_ID = 74 or vCONT_ID = 75 then
				Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT PARC_NumDocumento NumDocumento FROM Parcelas WITH (NOLOCK) WHERE PARC_NumDocumento LIKE '" & CDbl(vCPF) & "%' AND FPAG_ID = 3 AND Len(PARC_NumDocumento) = Len(" & CDbl(vCPF) & ") + 4 UNION SELECT BOAV_NumDocumento NumDocumento FROM Boletos_Avulsos WITH (NOLOCK) WHERE BOAV_NumDocumento LIKE '" & CDbl(vCPF) & "%' AND Len(BOAV_NumDocumento) = Len(" & CDbl(vCPF) & ") + 4 UNION SELECT BOGE_Numero NumDocumento FROM Boletos_Gerados WITH (NOLOCK) WHERE BOGE_Numero LIKE '" & CDbl(vCPF) & "%' AND Len(BOGE_Numero) = Len(" & CDbl(vCPF) & ") + 4) res")
				if IsNull(rsTemp("NumDocumento")) then
					vNumDoc = vCPF & "0001"
				else
					vNumDoc = rsTemp("NumDocumento")
				end if
			elseif vCONT_ID = 69 then
				Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT PARC_NumDocumento NumDocumento FROM Parcelas WITH (NOLOCK) WHERE PARC_NumDocumento LIKE '" & vCPF & "%' AND FPAG_ID = 3 AND Len(PARC_NumDocumento) = 7 UNION SELECT BOAV_NumDocumento NumDocumento FROM Boletos_Avulsos WITH (NOLOCK) WHERE BOAV_NumDocumento LIKE '" & vCPF & "%' AND Len(BOAV_NumDocumento) = 7 UNION SELECT BOGE_Numero NumDocumento FROM Boletos_Gerados WITH (NOLOCK) WHERE BOGE_Numero LIKE '" & vCPF & "%' AND Len(BOGE_Numero) = 7) res")
				if IsNull(rsTemp("NumDocumento")) then
					vNumDoc = vCPF & "001"
				else
					vNumDoc = rsTemp("NumDocumento")
				end if
			elseif vCONT_ID = 77 then
				vNumDoc = vCPF
			elseif vCONT_ID = 83 then
				if Month(rsParcelas("PARC_Vencimento")) = 10 then
					vNumDoc = "0"
				elseif Month(rsParcelas("PARC_Vencimento")) = 11 then
					vNumDoc = "5"
				elseif Month(rsParcelas("PARC_Vencimento")) = 12 then
					vNumDoc = "6"
				else
					vNumDoc = CSTR(Month(rsParcelas("PARC_Vencimento")))
				end if
				'Falta identificar o produto
				Set rsTemp = BD.Execute("SELECT BAND_Numero FROM Bandeiras b WITH (NOLOCK) JOIN Contratos c WITH (NOLOCK) ON b.BAND_ID = c.BAND_ID WHERE CTRA_ID = " & rsParcelas("CTRA_ID"))
				if rsTemp.EOF then
					vNumDoc = vNumDoc & "001" & Mid(vCPF, 7, 7)
				else
					vNumDoc = vNumDoc & Right("000" & rsTemp("BAND_Numero"), 3) & Mid(vCPF, 7, 7)
				end if
			elseif vCONT_ID = 97 then
				Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT PARC_NumDocumento NumDocumento FROM Parcelas WITH (NOLOCK) WHERE PARC_NumDocumento LIKE '" & vCPF & "%' AND FPAG_ID = 3 AND Len(PARC_NumDocumento) = 13 UNION SELECT BOAV_NumDocumento NumDocumento FROM Boletos_Avulsos WITH (NOLOCK) WHERE BOAV_NumDocumento LIKE '" & vCPF & "%' AND Len(BOAV_NumDocumento) = 13 UNION SELECT BOGE_Numero NumDocumento FROM Boletos_Gerados WITH (NOLOCK) WHERE BOGE_Numero LIKE '" & vCPF & "%' AND Len(BOGE_Numero) = 13) res")
				if IsNull(rsTemp("NumDocumento")) then
					vNumDoc = vCPF & "001"
				else
					vNumDoc = rsTemp("NumDocumento")
				end if
			else
				Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT PARC_NumDocumento NumDocumento FROM Parcelas WITH (NOLOCK) WHERE PARC_NumDocumento LIKE '" & vCPF & "%' AND FPAG_ID = 3 AND Len(PARC_NumDocumento) = 8 UNION SELECT BOAV_NumDocumento NumDocumento FROM Boletos_Avulsos WITH (NOLOCK) WHERE BOAV_NumDocumento LIKE '" & vCPF & "%' AND Len(BOAV_NumDocumento) = 8 UNION SELECT BOGE_Numero NumDocumento FROM Boletos_Gerados WITH (NOLOCK) WHERE BOGE_Numero LIKE '" & vCPF & "%' AND Len(BOGE_Numero) = 8) res")
				if IsNull(rsTemp("NumDocumento")) then
					vNumDoc = vCPF & "001"
				else
					vNumDoc = rsTemp("NumDocumento")
				end if
			end if
			BD.Execute "UPDATE Parcelas SET PARC_NumDocumento = " & vNumDoc & ", FPAG_ID = 3 WHERE PARC_ID = " & rsParcelas("PARC_ID")
			PreencheLOG "UPDATE Parcelas SET PARC_NumDocumento = " & vNumDoc & ", FPAG_ID = 3 WHERE PARC_ID = " & rsParcelas("PARC_ID"), vFilialDestino

			BD.Execute "INSERT INTO Boletos_Gerados(CTRA_ID, BOGE_Numero, PARC_ID, BOGE_TipoBoleto) VALUES(" & rsParcelas("CTRA_ID") & ", " & vNumDoc & ", " & rsParcelas("PARC_ID") & ", 'P')"
			PreencheLOG "INSERT INTO Boletos_Gerados(CTRA_ID, BOGE_Numero, PARC_ID, BOGE_TipoBoleto) VALUES(" & rsParcelas("CTRA_ID") & ", " & vNumDoc & ", " & rsParcelas("PARC_ID") & ", 'P')", vFilialDestino
		end if
		BD.Execute "UPDATE Parcelas SET PARC_BoletoEnviado = 1 WHERE PARC_ID = " & rsParcelas("PARC_ID")
		PreencheLOG "UPDATE Parcelas SET PARC_BoletoEnviado = 1 WHERE PARC_ID = " & rsParcelas("PARC_ID"), vFilialDestino

		rsParcelas.MoveNext
	Loop
	Set rsParcelas = BD.Execute("SELECT ACOR_ValorAcordo, CTRA_DataRecebimentoContrato, ACOR_Data, p.PARC_ID, p.ACOR_ID, PARC_Numero, res.Plano, res.ValorAcordo, Convert(char(10), BOAT_Vencimento, 103) PARC_Vencimento, BOAT_Valor PARC_ValorTotal, PARC_NumDocumento, c.CTRA_ID, CTRA_Numero, DEVE_CGCCPF, c.DEVE_ID, DEVE_Nome, EDEV_ID, BOAT_Validade BOAV_Validade, CART_ID, FNEG_ID, SCON_ID, BOAT_TaxaBoleto, CTRA_Conta FROM Parcelas p WITH (NOLOCK) JOIN Acordos a WITH(NOLOCK) ON p.ACOR_ID = a.ACOR_ID JOIN Contratos c WITH (NOLOCK) ON a.CTRA_ID = c.CTRA_ID JOIN Devedores d WITH (NOLOCK) ON c.DEVE_ID = d.DEVE_ID JOIN (SELECT ACOR_ID, COUNT(*) Plano, SUM(PARC_ValorTotal) ValorAcordo FROM Parcelas WITH (NOLOCK) GROUP BY ACOR_ID) res ON p.ACOR_ID = res.ACOR_ID JOIN Boletos_Atraso_Temp bp WITH (NOLOCK) ON bp.PARC_ID = p.PARC_ID WHERE bp.FUNC_ID = " & Session("FUNC_ID"))
elseif Request.QueryString("avulso") <> "" then
	vTpBoleto = "A"
	'Avulso Acumulado
	Set rsParcelas = BD.Execute("SELECT * FROM Boletos_Avulsos_Temp WHERE FUNC_ID = " & Session("FUNC_ID"))
	Do While Not rsParcelas.EOF
		Set rsFilial = BD.Execute("SELECT ca.CONT_ID, FILI_ID FROM Carteiras ca WITH (NOLOCK) JOIN Contratos c WITH (NOLOCK) ON c.CART_ID = ca.CART_ID WHERE CTRA_ID = " & rsParcelas("CTRA_ID"))
		vFilialDestino = rsFilial("FILI_ID")


		vCONT_ID = rsFilial("CONT_ID")
		
		rsFilial.Close
		Set rsFilial = Nothing

		Set rsCPF = BD.Execute("SELECT DEVE_CGCCPF, CTRA_Numero FROM Devedores d WITH (NOLOCK) JOIN Contratos c WITH (NOLOCK) ON c.DEVE_ID = d.DEVE_ID WHERE CTRA_ID = " & rsParcelas("CTRA_ID"))
		if vCONT_ID = 74 or vCONT_ID = 75 then
			if len(rsCPF("DEVE_CGCCPF")) = 14 then
				vCPF = "0" & Mid(rsCPF("DEVE_CGCCPF"), 1, 8)
			else
				vCPF = Mid(rsCPF("DEVE_CGCCPF"), 1, 9)
			end if
		elseif vCONT_ID = 69 then
			vCPF = CDbl(rsCPF("DEVE_CGCCPF"))
			vCPF = Mid(vCPF & "0000", 1, 4)
		elseif vCONT_ID = 77 then
			vCPF = rsCPF("CTRA_Numero")
			vCPF = Right("00000000000000" & vCPF, 14)
		elseif vCONT_ID = 83 then
			vCPF = rsCPF("CTRA_Numero")
		elseif vCONT_ID = 97 then
			vCPF = CDbl(rsCPF("DEVE_CGCCPF"))
			vCPF = "1200905" & Mid(vCPF & "000", 1, 3)
		else
			vCPF = CDbl(rsCPF("DEVE_CGCCPF"))
			vCPF = Mid(vCPF & "00000", 1, 5)
		end if

		Set rsBanco = BD.Execute("SELECT CONT_InicioSeqBoleto FROM Contratante WITH (NOLOCK) WHERE CONT_ID = " & vCONT_ID)
		if not isnull(rsBanco("CONT_InicioSeqBoleto")) then
			vInicioSeqBoleto = rsBanco("CONT_InicioSeqBoleto") 'Substituir por campo no banco de dados
		else
			vInicioSeqBoleto = 0
		end if

		if vCONT_ID = 11 then
			Set rsNumAco = BD.Execute("SELECT CTRA_NumeroAcordoContratante FROM Contratos c WITH (NOLOCK) WHERE CTRA_ID = " & rsParcelas("CTRA_ID"))
			if isnull(rsNumAco("CTRA_NumeroAcordoContratante")) then
				Set rsNumAco = BD.Execute("SELECT MAX(CTRA_NumeroAcordoContratante) + 1 NumAcordo FROM Contratos c WITH (NOLOCK) JOIN Carteiras ca WITH (NOLOCK) ON c.CART_ID = ca.CART_ID WHERE CONT_ID = " & vCONT_ID)
				BD.Execute "UPDATE Contratos SET CTRA_NumeroAcordoContratante = " & rsNumAco("NumAcordo") & " WHERE CTRA_ID = " & rsParcelas("CTRA_ID") 
				PreencheLOG "UPDATE Contratos SET CTRA_NumeroAcordoContratante = " & rsNumAco("NumAcordo") & " WHERE CTRA_ID = " & rsParcelas("CTRA_ID") , vFilialDestino
			end if
			Set rsNumAco = BD.Execute("SELECT CTRA_NumeroAcordoContratante FROM Contratos c WITH (NOLOCK) WHERE CTRA_ID = " & rsParcelas("CTRA_ID"))
            if not isnull(rsNumAco("CTRA_NumeroAcordoContratante")) then
				vAcordoNumero = Right("000000" & rsNumAco("CTRA_NumeroAcordoContratante"), 6)
			end if
			vNumDoc = vAcordoNumero & "001" & Right("00" & rsParcelas("boat_qtdparcelas"), 2)
		elseif vCONT_ID = 14 then
			Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT BOGE_Numero NumDocumento FROM Boletos_Gerados b WITH (NOLOCK) JOIN Contratos c WITH (NOLOCK) ON b.CTRA_ID = c.CTRA_ID JOIN Carteiras ca WITH (NOLOCK) ON c.CART_ID = ca.CART_ID WHERE CONT_ID = " & vCONT_ID & " AND BOGE_Numero >= " & vInicioSeqBoleto & " AND Len(BOGE_Numero) = 9) res")
			if IsNull(rsTemp("NumDocumento")) then
				vNumDoc = vInicioSeqBoleto
			else
				vNumDoc = rsTemp("NumDocumento")
			end if
		elseif vCONT_ID = 74 or vCONT_ID = 75 then
			Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT PARC_NumDocumento NumDocumento FROM Parcelas WITH (NOLOCK) WHERE PARC_NumDocumento LIKE '" & CDbl(vCPF) & "%' AND FPAG_ID = 3 AND Len(PARC_NumDocumento) = Len(" & CDbl(vCPF) & ") + 4 UNION SELECT BOAV_NumDocumento NumDocumento FROM Boletos_Avulsos WITH (NOLOCK) WHERE BOAV_NumDocumento LIKE '" & CDbl(vCPF) & "%' AND Len(BOAV_NumDocumento) = Len(" & CDbl(vCPF) & ") + 4 UNION SELECT BOGE_Numero NumDocumento FROM Boletos_Gerados WITH (NOLOCK) WHERE BOGE_Numero LIKE '" & CDbl(vCPF) & "%' AND Len(BOGE_Numero) = Len(" & CDbl(vCPF) & ") + 4) res")
			if IsNull(rsTemp("NumDocumento")) then
				vNumDoc = vCPF & "0001"
			else
				vNumDoc = rsTemp("NumDocumento")
			end if
		elseif vCONT_ID = 69 then
			Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT PARC_NumDocumento NumDocumento FROM Parcelas WITH (NOLOCK) WHERE PARC_NumDocumento LIKE '" & vCPF & "%' AND FPAG_ID = 3 AND Len(PARC_NumDocumento) = 7 UNION SELECT BOAV_NumDocumento NumDocumento FROM Boletos_Avulsos WITH (NOLOCK) WHERE BOAV_NumDocumento LIKE '" & vCPF & "%' AND Len(BOAV_NumDocumento) = 7 UNION SELECT BOGE_Numero NumDocumento FROM Boletos_Gerados WITH (NOLOCK) WHERE BOGE_Numero LIKE '" & vCPF & "%' AND Len(BOGE_Numero) = 7) res")
			if IsNull(rsTemp("NumDocumento")) then
				vNumDoc = vCPF & "001"
			else
				vNumDoc = rsTemp("NumDocumento")
			end if
		elseif vCONT_ID = 77 then
			vNumDoc = vCPF
		elseif vCONT_ID = 83 then
			if Month(rsParcelas("BOAT_Vencimento")) = 10 then
				vNumDoc = "0"
			elseif Month(rsParcelas("BOAT_Vencimento")) = 11 then
				vNumDoc = "5"
			elseif Month(rsParcelas("BOAT_Vencimento")) = 12 then
				vNumDoc = "6"
			else
				vNumDoc = CSTR(Month(rsParcelas("BOAT_Vencimento")))
			end if
			Set rsTemp = BD.Execute("SELECT BAND_Numero FROM Bandeiras b WITH (NOLOCK) JOIN Contratos c WITH (NOLOCK) ON b.BAND_ID = c.BAND_ID WHERE CTRA_ID = " & rsParcelas("CTRA_ID"))
			if rsTemp.EOF then
				vNumDoc = vNumDoc & "001" & Mid(vCPF, 7, 7)
			else
				vNumDoc = vNumDoc & Right("000" & rsTemp("BAND_Numero"), 3) & Mid(vCPF, 7, 7)
			end if
		elseif vCONT_ID = 97 then
			Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT PARC_NumDocumento NumDocumento FROM Parcelas WITH (NOLOCK) WHERE PARC_NumDocumento LIKE '" & vCPF & "%' AND FPAG_ID = 3 AND Len(PARC_NumDocumento) = 13 UNION SELECT BOAV_NumDocumento NumDocumento FROM Boletos_Avulsos WITH (NOLOCK) WHERE BOAV_NumDocumento LIKE '" & vCPF & "%' AND Len(BOAV_NumDocumento) = 13 UNION SELECT BOGE_Numero NumDocumento FROM Boletos_Gerados WITH (NOLOCK) WHERE BOGE_Numero LIKE '" & vCPF & "%' AND Len(BOGE_Numero) = 13) res")
			if IsNull(rsTemp("NumDocumento")) then
				vNumDoc = vCPF & "001"
			else
				vNumDoc = rsTemp("NumDocumento")
			end if
		else
			Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT PARC_NumDocumento NumDocumento FROM Parcelas WITH (NOLOCK) WHERE PARC_NumDocumento LIKE '" & vCPF & "%' AND FPAG_ID = 3 AND Len(PARC_NumDocumento) = 8 UNION SELECT BOAV_NumDocumento NumDocumento FROM Boletos_Avulsos WITH (NOLOCK) WHERE BOAV_NumDocumento LIKE '" & vCPF & "%' AND Len(BOAV_NumDocumento) = 8 UNION SELECT BOGE_Numero NumDocumento FROM Boletos_Gerados WITH (NOLOCK) WHERE BOGE_Numero LIKE '" & vCPF & "%' AND Len(BOGE_Numero) = 8) res")
			'Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT PARC_NumDocumento NumDocumento FROM Parcelas WITH (NOLOCK) WHERE PARC_NumDocumento LIKE '" & vCPF & "%' AND FPAG_ID = 3 AND Len(PARC_NumDocumento) = 8 UNION SELECT BOAV_NumDocumento NumDocumento FROM Boletos_Avulsos WITH (NOLOCK) WHERE BOAV_NumDocumento LIKE '" & vCPF & "%' AND Len(BOAV_NumDocumento) = 8) res")
			'Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT PARC_NumDocumento NumDocumento FROM Parcelas WITH (NOLOCK) WHERE PARC_NumDocumento LIKE '" & Session("FILI_ID") & "%' AND FPAG_ID = 3 AND Len(PARC_NumDocumento) >= 8 AND Len(PARC_NumDocumento) <=9 UNION SELECT BOAV_NumDocumento NumDocumento FROM Boletos_Avulsos WITH (NOLOCK) WHERE BOAV_NumDocumento LIKE '" & Session("FILI_ID") & "%' AND Len(BOAV_NumDocumento) >= 8 AND Len(BOAV_NumDocumento) <=9) res")
			'Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT PARC_NumDocumento NumDocumento FROM Parcelas WITH (NOLOCK) WHERE PARC_NumDocumento LIKE '" & Session("FILI_ID") & "%' AND FPAG_ID = 3 AND Len(PARC_NumDocumento) >= 7 UNION SELECT BOAV_NumDocumento NumDocumento FROM Boletos_Avulsos WITH (NOLOCK) WHERE BOAV_NumDocumento LIKE '" & Session("FILI_ID") & "%' AND Len(BOAV_NumDocumento) >= 7) res")
			if IsNull(rsTemp("NumDocumento")) then
				vNumDoc = vCPF & "001"
			else
				vNumDoc = rsTemp("NumDocumento")
			end if
		end if

		vID = RetornaProximoNumero("Boletos_Avulsos", "BOAV_ID", Session("FILI_ID"))
		BD.Execute "INSERT INTO Boletos_Avulsos Values(" & vID & ", " & rsParcelas("ctra_id") & ", " & rsParcelas("edev_id") & ", '" & vNumDoc & "', 'A', " & Replace(Replace(rsParcelas("boat_valor"),".",""),",",".") & ", '" & rsParcelas("boat_vencimento") & "', " & rsParcelas("boat_qtdparcelas") & ", '" & rsParcelas("boat_validade") & "', " & Replace(Replace(rsParcelas("boat_taxaboleto"),".",""),",",".") & ", NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL)"
		PreencheLOG "INSERT INTO Boletos_Avulsos Values(" & vID & ", " & rsParcelas("ctra_id") & ", " & rsParcelas("edev_id") & ", '" & vNumDoc & "', 'A', " & Replace(Replace(rsParcelas("boat_valor"),".",""),",",".") & ", '" & rsParcelas("boat_vencimento") & "', " & rsParcelas("boat_qtdparcelas") & ", '" & rsParcelas("boat_validade") & "', " & Replace(Replace(rsParcelas("boat_taxaboleto"),".",""),",",".") & ", NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL)", vFilialDestino

		BD.Execute "INSERT INTO Boletos_Gerados(CTRA_ID, BOGE_Numero, BOGE_TipoBoleto) VALUES(" & rsParcelas("CTRA_ID") & ", " & vNumDoc & ", 'A')"
		PreencheLOG "INSERT INTO Boletos_Gerados(CTRA_ID, BOGE_Numero, BOGE_TipoBoleto) VALUES(" & rsParcelas("CTRA_ID") & ", " & vNumDoc & ", 'A')", vFilialDestino

		vBoletos = vBoletos & vID & ","

		rsParcelas.MoveNext
	Loop
	vBoletos = left(vBoletos, len(vBoletos) - 1)
	BD.Execute("DELETE FROM Boletos_Avulsos_Temp WHERE FUNC_ID = " & Session("FUNC_ID"))
	Set rsParcelas = BD.Execute("SELECT 0 PARC_ID, b.CTRA_ID, d.DEVE_ID, DEVE_Nome, DEVE_CGCCPF, BOAV_NumDocumento PARC_NumDocumento, EDEV_ID, CTRA_Numero, b.BOAV_QtdParcelas Plano, BOAV_TaxaBoleto, BOAV_Validade, BOAV_Validade, BOAV_Vencimento, BOAV_Valor, CART_ID, CTRA_VencDebito, SCON_ID, CTRA_Conta FROM Boletos_Avulsos b WITH (NOLOCK) JOIN Contratos c WITH (NOLOCK) ON b.CTRA_ID = c.CTRA_ID JOIN Devedores d WITH (NOLOCK) ON c.DEVE_ID = d.DEVE_ID WHERE BOAV_ID IN (" & vBoletos & ") AND ACMD_ID IS NULL")
elseif vTipoBoleto = 1 then
	vTpBoleto = "P"

	Set rsFilial = BD.Execute("SELECT ca.CONT_ID, FILI_ID FROM Carteiras ca WITH (NOLOCK) JOIN Contratos c WITH (NOLOCK) ON c.CART_ID = ca.CART_ID WHERE CTRA_ID = " & vCTRA_ID)
	vFilialDestino = rsFilial("FILI_ID")
	
	vCONT_ID = rsFilial("CONT_ID")

	rsFilial.Close
	Set rsFilial = Nothing
	'Laser Acordo

	Set rsCPF = BD.Execute("SELECT DEVE_CGCCPF, CTRA_Numero FROM Devedores d WITH (NOLOCK) JOIN Contratos c WITH (NOLOCK) ON c.DEVE_ID = d.DEVE_ID WHERE CTRA_ID = " & vCTRA_ID)
	if vCONT_ID = 74 or vCONT_ID = 75 then
		if len(rsCPF("DEVE_CGCCPF")) = 14 then
			vCPF = "0" & Mid(rsCPF("DEVE_CGCCPF"), 1, 8)
		else
			vCPF = Mid(rsCPF("DEVE_CGCCPF"), 1, 9)
		end if
	elseif vCONT_ID = 69 then
		vCPF = CDbl(rsCPF("DEVE_CGCCPF"))
		vCPF = Mid(vCPF & "0000", 1, 4)
	elseif vCONT_ID = 77 then
		vCPF = rsCPF("CTRA_Numero")
		vCPF = Right("00000000000000" & vCPF, 14)
	elseif vCONT_ID = 83 then
		vCPF = rsCPF("CTRA_Numero")
	elseif vCONT_ID = 97 then
		vCPF = CDbl(rsCPF("DEVE_CGCCPF"))
		vCPF = "1200905" & Mid(vCPF & "000", 1, 3)
	else
		vCPF = CDbl(rsCPF("DEVE_CGCCPF"))
		vCPF = Mid(vCPF & "00000", 1, 5)
	end if

	vParcelas = ""
	For Each Key In Request.Form
		if mid(key, 1, 4) = "cbox" then
			vParcelas = vParcelas & Request.Form(Key) & ","
		end if
	Next
	vParcelas = left(vParcelas, len(vParcelas) - 1)

	Set rsBanco = BD.Execute("SELECT CONT_InicioSeqBoleto FROM Contratante WITH (NOLOCK) WHERE CONT_ID = " & vCONT_ID)
	if not isnull(rsBanco("CONT_InicioSeqBoleto")) then
		vInicioSeqBoleto = rsBanco("CONT_InicioSeqBoleto") 'Substituir por campo no banco de dados
	else
		vInicioSeqBoleto = 0
	end if

	Set rsParcelas = BD.Execute("SELECT p.ACOR_ID, PARC_ID, PARC_NumDocumento, FPAG_ID, PARC_Numero, Qtd Plano, CTRA_ID, PARC_Vencimento FROM Parcelas p WITH (NOLOCK) JOIN Acordos a WITH (NOLOCK) ON a.ACOR_ID = p.ACOR_ID JOIN (SELECT ACOR_ID, MAX(PARC_Numero) Qtd FROM Parcelas WITH (NOLOCK) GROUP BY ACOR_ID) r ON r.ACOR_ID = p.ACOR_ID WHERE PARC_ID IN (" & vParcelas & ") ORDER BY PARC_Numero")
	Do While Not rsParcelas.EOF
		'if (vCONT_ID = 11 and Len(rsParcelas("PARC_NumDocumento")) < 9) or ((rsParcelas("FPAG_ID") = 3 and (IsNull(rsParcelas("PARC_NumDocumento")) or rsParcelas("PARC_NumDocumento") = "" or rsParcelas("PARC_NumDocumento") = "0")) or (rsParcelas("FPAG_ID") <> 3) or IsNull(rsParcelas("FPAG_ID"))) then
		if (vCONT_ID = 11) or ((rsParcelas("FPAG_ID") = 3 and (IsNull(rsParcelas("PARC_NumDocumento")) or rsParcelas("PARC_NumDocumento") = "" or rsParcelas("PARC_NumDocumento") = "0")) or (rsParcelas("FPAG_ID") <> 3) or IsNull(rsParcelas("FPAG_ID"))) then
			if vCONT_ID = 11 then
				Set rsNumAco = BD.Execute("SELECT CTRA_NumeroAcordoContratante FROM Contratos c WITH (NOLOCK) WHERE CTRA_ID = " & rsParcelas("CTRA_ID"))
                if not isnull(rsNumAco("CTRA_NumeroAcordoContratante")) then
					vAcordoNumero = Right("000000" & rsNumAco("CTRA_NumeroAcordoContratante"), 6)
				else
					If Len(rsParcelas("ACOR_ID")) > 5 Then
						vAcordoNumero = Right("000000" & Mid(rsParcelas("ACOR_ID"), 1, Len(rsParcelas("ACOR_ID")) - 3), 6)
					Else
						vAcordoNumero = Right("000000" & rsParcelas("ACOR_ID"), 6)
					End If
				end if
				vNumDoc = vAcordoNumero & "0" & Right("00" & rsParcelas("PARC_Numero"), 2) & Right("00" & rsParcelas("Plano"), 2)
			elseif vCONT_ID = 14 then
				Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT BOGE_Numero NumDocumento FROM Boletos_Gerados b WITH (NOLOCK) JOIN Contratos c WITH (NOLOCK) ON b.CTRA_ID = c.CTRA_ID JOIN Carteiras ca WITH (NOLOCK) ON c.CART_ID = ca.CART_ID WHERE CONT_ID = " & vCONT_ID & " AND BOGE_Numero >= " & vInicioSeqBoleto & " AND Len(BOGE_Numero) = 9) res")
				if IsNull(rsTemp("NumDocumento")) then
					vNumDoc = vInicioSeqBoleto
				else
					vNumDoc = rsTemp("NumDocumento")
				end if
			elseif vCONT_ID = 74 or vCONT_ID = 75 then
				Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT PARC_NumDocumento NumDocumento FROM Parcelas WITH (NOLOCK) WHERE PARC_NumDocumento LIKE '" & CDbl(vCPF) & "%' AND FPAG_ID = 3 AND Len(PARC_NumDocumento) = Len(" & CDbl(vCPF) & ") + 4 UNION SELECT BOAV_NumDocumento NumDocumento FROM Boletos_Avulsos WITH (NOLOCK) WHERE BOAV_NumDocumento LIKE '" & CDbl(vCPF) & "%' AND Len(BOAV_NumDocumento) = Len(" & CDbl(vCPF) & ") + 4 UNION SELECT BOGE_Numero NumDocumento FROM Boletos_Gerados WITH (NOLOCK) WHERE BOGE_Numero LIKE '" & CDbl(vCPF) & "%' AND Len(BOGE_Numero) = Len(" & CDbl(vCPF) & ") + 4) res")
				if IsNull(rsTemp("NumDocumento")) then
					vNumDoc = vCPF & "0001"
				else
					vNumDoc = rsTemp("NumDocumento")
				end if
			elseif vCONT_ID = 69 then
				Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT PARC_NumDocumento NumDocumento FROM Parcelas WITH (NOLOCK) WHERE PARC_NumDocumento LIKE '" & vCPF & "%' AND FPAG_ID = 3 AND Len(PARC_NumDocumento) = 7 UNION SELECT BOAV_NumDocumento NumDocumento FROM Boletos_Avulsos WITH (NOLOCK) WHERE BOAV_NumDocumento LIKE '" & vCPF & "%' AND Len(BOAV_NumDocumento) = 7 UNION SELECT BOGE_Numero NumDocumento FROM Boletos_Gerados WITH (NOLOCK) WHERE BOGE_Numero LIKE '" & vCPF & "%' AND Len(BOGE_Numero) = 7) res")
				if IsNull(rsTemp("NumDocumento")) then
					vNumDoc = vCPF & "001"
				else
					vNumDoc = rsTemp("NumDocumento")
				end if
			elseif vCONT_ID = 77 then
				vNumDoc = vCPF
			elseif vCONT_ID = 83 then
				if Month(rsParcelas("PARC_Vencimento")) = 10 then
					vNumDoc = "0"
				elseif Month(rsParcelas("PARC_Vencimento")) = 11 then
					vNumDoc = "5"
				elseif Month(rsParcelas("PARC_Vencimento")) = 12 then
					vNumDoc = "6"
				else
					vNumDoc = CSTR(Month(rsParcelas("PARC_Vencimento")))
				end if
				Set rsTemp = BD.Execute("SELECT BAND_Numero FROM Bandeiras b WITH (NOLOCK) JOIN Contratos c WITH (NOLOCK) ON b.BAND_ID = c.BAND_ID WHERE CTRA_ID = " & rsParcelas("CTRA_ID"))
				if rsTemp.EOF then
					vNumDoc = vNumDoc & "001" & Mid(vCPF, 7, 7)
				else
					vNumDoc = vNumDoc & Right("000" & rsTemp("BAND_Numero"), 3) & Mid(vCPF, 7, 7)
				end if
			elseif vCONT_ID = 97 then
				Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT PARC_NumDocumento NumDocumento FROM Parcelas WITH (NOLOCK) WHERE PARC_NumDocumento LIKE '" & vCPF & "%' AND FPAG_ID = 3 AND Len(PARC_NumDocumento) = 13 UNION SELECT BOAV_NumDocumento NumDocumento FROM Boletos_Avulsos WITH (NOLOCK) WHERE BOAV_NumDocumento LIKE '" & vCPF & "%' AND Len(BOAV_NumDocumento) = 13 UNION SELECT BOGE_Numero NumDocumento FROM Boletos_Gerados WITH (NOLOCK) WHERE BOGE_Numero LIKE '" & vCPF & "%' AND Len(BOGE_Numero) = 13) res")
				if IsNull(rsTemp("NumDocumento")) then
					vNumDoc = vCPF & "001"
				else
					vNumDoc = rsTemp("NumDocumento")
				end if
			else
				Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT PARC_NumDocumento NumDocumento FROM Parcelas WITH (NOLOCK) WHERE PARC_NumDocumento LIKE '" & vCPF & "%' AND FPAG_ID = 3 AND Len(PARC_NumDocumento) = 8 UNION SELECT BOAV_NumDocumento NumDocumento FROM Boletos_Avulsos WITH (NOLOCK) WHERE BOAV_NumDocumento LIKE '" & vCPF & "%' AND Len(BOAV_NumDocumento) = 8 UNION SELECT BOGE_Numero NumDocumento FROM Boletos_Gerados WITH (NOLOCK) WHERE BOGE_Numero LIKE '" & vCPF & "%' AND Len(BOGE_Numero) = 8) res")
				'Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT PARC_NumDocumento NumDocumento FROM Parcelas WITH (NOLOCK) WHERE PARC_NumDocumento LIKE '" & vCPF & "%' AND FPAG_ID = 3 AND Len(PARC_NumDocumento) = 8 UNION SELECT BOAV_NumDocumento NumDocumento FROM Boletos_Avulsos WITH (NOLOCK) WHERE BOAV_NumDocumento LIKE '" & vCPF & "%' AND Len(BOAV_NumDocumento) = 8) res")
				'Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT PARC_NumDocumento NumDocumento FROM Parcelas WITH (NOLOCK) WHERE PARC_NumDocumento LIKE '" & Session("FILI_ID") & "%' AND FPAG_ID = 3 AND Len(PARC_NumDocumento) >= 8 AND Len(PARC_NumDocumento) <= 9 UNION SELECT BOAV_NumDocumento NumDocumento FROM Boletos_Avulsos WITH (NOLOCK) WHERE BOAV_NumDocumento LIKE '" & Session("FILI_ID") & "%' AND Len(BOAV_NumDocumento) >= 8 AND Len(BOAV_NumDocumento) <= 9) res")
				'Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT PARC_NumDocumento NumDocumento FROM Parcelas WITH (NOLOCK) WHERE PARC_NumDocumento LIKE '" & Session("FILI_ID") & "%' AND FPAG_ID = 3 AND Len(PARC_NumDocumento) >= 7 UNION SELECT BOAV_NumDocumento NumDocumento FROM Boletos_Avulsos WITH (NOLOCK) WHERE BOAV_NumDocumento LIKE '" & Session("FILI_ID") & "%' AND Len(BOAV_NumDocumento) >= 7) res")
				if IsNull(rsTemp("NumDocumento")) then
					vNumDoc = vCPF & "001"
				else
					vNumDoc = rsTemp("NumDocumento")
				end if
			end if
			BD.Execute "UPDATE Parcelas SET PARC_NumDocumento = " & vNumDoc & ", FPAG_ID = 3 WHERE PARC_ID = " & rsParcelas("PARC_ID")
			PreencheLOG "UPDATE Parcelas SET PARC_NumDocumento = " & vNumDoc & ", FPAG_ID = 3 WHERE PARC_ID = " & rsParcelas("PARC_ID"), vFilialDestino

			BD.Execute "INSERT INTO Boletos_Gerados(CTRA_ID, BOGE_Numero, PARC_ID, BOGE_TipoBoleto) VALUES(" & vCTRA_ID & ", " & vNumDoc & ", " & rsParcelas("PARC_ID") & ", 'P')"
			PreencheLOG "INSERT INTO Boletos_Gerados(CTRA_ID, BOGE_Numero, PARC_ID, BOGE_TipoBoleto) VALUES(" & vCTRA_ID & ", " & vNumDoc & ", " & rsParcelas("PARC_ID") & ", 'P')", vFilialDestino
		end if
		BD.Execute "UPDATE Parcelas SET PARC_BoletoEnviado = 1 WHERE PARC_ID = " & rsParcelas("PARC_ID")
		PreencheLOG "UPDATE Parcelas SET PARC_BoletoEnviado = 1 WHERE PARC_ID = " & rsParcelas("PARC_ID"), vFilialDestino

		rsParcelas.MoveNext
	Loop

	Set rsParcelas = BD.Execute("SELECT ACOR_ValorAcordo, CTRA_DataRecebimentoContrato, ACOR_Data, PARC_ID, p.ACOR_ID, PARC_Numero, res.Plano, res.ValorAcordo, Convert(char(10), PARC_Vencimento, 103) PARC_Vencimento, PARC_ValorTotal, PARC_NumDocumento, c.CTRA_ID, CTRA_Numero, DEVE_CGCCPF, c.DEVE_ID, DEVE_Nome, " & vE & " EDEV_ID, 0 BOAV_Validade, CART_ID, FNEG_ID, SCON_ID, CTRA_Conta FROM Parcelas p WITH (NOLOCK) JOIN Acordos a WITH(NOLOCK) ON p.ACOR_ID = a.ACOR_ID JOIN Contratos c WITH (NOLOCK) ON a.CTRA_ID = c.CTRA_ID JOIN Devedores d WITH (NOLOCK) ON c.DEVE_ID = d.DEVE_ID JOIN (SELECT ACOR_ID, MAX(PARC_Numero) Plano, SUM(PARC_ValorTotal) ValorAcordo FROM Parcelas WITH (NOLOCK) GROUP BY ACOR_ID) res ON p.ACOR_ID = res.ACOR_ID WHERE PARC_ID IN (" & vParcelas & ") ORDER BY PARC_Numero")
elseif vTipoBoleto = 4 then
	vTpBoleto = "T"

	Set rsFilial = BD.Execute("SELECT ca.CONT_ID, FILI_ID FROM Carteiras ca WITH (NOLOCK) JOIN Contratos c WITH (NOLOCK) ON c.CART_ID = ca.CART_ID WHERE CTRA_ID = " & vCTRA_ID)
	vFilialDestino = rsFilial("FILI_ID")

	vCONT_ID = rsFilial("CONT_ID")
	
	
	rsFilial.Close
	Set rsFilial = Nothing
	'Laser Acordo

	Set rsCPF = BD.Execute("SELECT DEVE_CGCCPF, CTRA_Numero FROM Devedores d WITH (NOLOCK) JOIN Contratos c WITH (NOLOCK) ON c.DEVE_ID = d.DEVE_ID WHERE CTRA_ID = " & vCTRA_ID)
	if vCONT_ID = 74 or vCONT_ID = 75 then
		if len(rsCPF("DEVE_CGCCPF")) = 14 then
			vCPF = "0" & Mid(rsCPF("DEVE_CGCCPF"), 1, 8)
		else
			vCPF = Mid(rsCPF("DEVE_CGCCPF"), 1, 9)
		end if
	elseif vCONT_ID = 69 then
		vCPF = CDbl(rsCPF("DEVE_CGCCPF"))
		vCPF = Mid(vCPF & "0000", 1, 4)
	elseif vCONT_ID = 77 then
		vCPF = rsCPF("CTRA_Numero")
		vCPF = Right("00000000000000" & vCPF, 14)
	elseif vCONT_ID = 97 then
		vCPF = CDbl(rsCPF("DEVE_CGCCPF"))
		vCPF = "1200905" & Mid(vCPF & "000", 1, 3)
	else
		vCPF = CDbl(rsCPF("DEVE_CGCCPF"))
		vCPF = Mid(vCPF & "00000", 1, 5)
	end if

	vTitulos = ""
	For Each Key In Request.Form
		if mid(key, 1, 4) = "cbox" then
			vTitulos = vTitulos & Request.Form(Key) & ","
		end if
	Next
	vTitulos = left(vTitulos, len(vTitulos) - 1)
	
	if Request.Form("rdCobraBoleto") = "Sim" then
		vValTx = Replace(Request.Form("ValorTaxaBoleto"), ",", ".")
	else
		vValTx = 0
	end if
	
	Set rsBanco = BD.Execute("SELECT CONT_InicioSeqBoleto FROM Contratante WITH (NOLOCK) WHERE CONT_ID = " & vCONT_ID)
	if not isnull(rsBanco("CONT_InicioSeqBoleto")) then
		vInicioSeqBoleto = rsBanco("CONT_InicioSeqBoleto") 'Substituir por campo no banco de dados
	else
		vInicioSeqBoleto = 0
	end if

	Set rsParcelas = BD.Execute("SELECT TRAN_ID, TRAN_NumTitulo, CTRA_ID FROM Transacoes t WITH (NOLOCK) WHERE TRAN_ID IN (" & vTitulos & ")")
	Do While Not rsParcelas.EOF
		'if vCONT_ID = 11 then
		'	Set rsNumAco = BD.Execute("SELECT CTRA_NumeroAcordoContratante FROM Contratos c WITH (NOLOCK) WHERE CTRA_ID = " & rsParcelas("CTRA_ID"))
		'	if isnull(rsNumAco("CTRA_NumeroAcordoContratante")) then
		'		Set rsNumAco = BD.Execute("SELECT MAX(CTRA_NumeroAcordoContratante) + 1 NumAcordo FROM Contratos c WITH (NOLOCK) JOIN Carteiras ca WITH (NOLOCK) ON c.CART_ID = ca.CART_ID WHERE CONT_ID = " & vCONT_ID)
		'		BD.Execute "UPDATE Contratos SET CTRA_NumeroAcordoContratante = " & rsNumAco("NumAcordo") & " WHERE CTRA_ID = " & rsParcelas("CTRA_ID") 
		'		PreencheLOG "UPDATE Contratos SET CTRA_NumeroAcordoContratante = " & rsNumAco("NumAcordo") & " WHERE CTRA_ID = " & rsParcelas("CTRA_ID") , vFilialDestino
		'	end if
		'	Set rsNumAco = BD.Execute("SELECT CTRA_NumeroAcordoContratante FROM Contratos c WITH (NOLOCK) WHERE CTRA_ID = " & rsParcelas("CTRA_ID"))
        '   if not isnull(rsNumAco("CTRA_NumeroAcordoContratante")) then
		'		vAcordoNumero = Right("000000" & rsNumAco("CTRA_NumeroAcordoContratante"), 6)
		'	end if
		'	vNumDoc = vAcordoNumero & "001" & Right("00" & rsParcelas("boat_qtdparcelas"), 2)
		'elseif vCONT_ID = 14 then
		if vCONT_ID = 14 then
			Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT BOGE_Numero NumDocumento FROM Boletos_Gerados b WITH (NOLOCK) JOIN Contratos c WITH (NOLOCK) ON b.CTRA_ID = c.CTRA_ID JOIN Carteiras ca WITH (NOLOCK) ON c.CART_ID = ca.CART_ID WHERE CONT_ID = " & vCONT_ID & " AND BOGE_Numero >= " & vInicioSeqBoleto & " AND Len(BOGE_Numero) = 9) res")
			if IsNull(rsTemp("NumDocumento")) then
				vNumDoc = vInicioSeqBoleto
			else
				vNumDoc = rsTemp("NumDocumento")
			end if
		elseif vCONT_ID = 74 or vCONT_ID = 75 then
			Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT PARC_NumDocumento NumDocumento FROM Parcelas WITH (NOLOCK) WHERE PARC_NumDocumento LIKE '" & CDbl(vCPF) & "%' AND FPAG_ID = 3 AND Len(PARC_NumDocumento) = Len(" & CDbl(vCPF) & ") + 4 UNION SELECT BOAV_NumDocumento NumDocumento FROM Boletos_Avulsos WITH (NOLOCK) WHERE BOAV_NumDocumento LIKE '" & CDbl(vCPF) & "%' AND Len(BOAV_NumDocumento) = Len(" & CDbl(vCPF) & ") + 4 UNION SELECT BOGE_Numero NumDocumento FROM Boletos_Gerados WITH (NOLOCK) WHERE BOGE_Numero LIKE '" & CDbl(vCPF) & "%' AND Len(BOGE_Numero) = Len(" & CDbl(vCPF) & ") + 4) res")
			if IsNull(rsTemp("NumDocumento")) then
				vNumDoc = vCPF & "0001"
			else
				vNumDoc = rsTemp("NumDocumento")
			end if
		elseif vCONT_ID = 69 then
			Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT PARC_NumDocumento NumDocumento FROM Parcelas WITH (NOLOCK) WHERE PARC_NumDocumento LIKE '" & vCPF & "%' AND FPAG_ID = 3 AND Len(PARC_NumDocumento) = 7 UNION SELECT BOAV_NumDocumento NumDocumento FROM Boletos_Avulsos WITH (NOLOCK) WHERE BOAV_NumDocumento LIKE '" & vCPF & "%' AND Len(BOAV_NumDocumento) = 7 UNION SELECT BOGE_Numero NumDocumento FROM Boletos_Gerados WITH (NOLOCK) WHERE BOGE_Numero LIKE '" & vCPF & "%' AND Len(BOGE_Numero) = 7) res")
			if IsNull(rsTemp("NumDocumento")) then
				vNumDoc = vCPF & "001"
			else
				vNumDoc = rsTemp("NumDocumento")
			end if
		elseif vCONT_ID = 77 then
			vNumDoc = vCPF
		elseif vCONT_ID = 97 then
			Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT PARC_NumDocumento NumDocumento FROM Parcelas WITH (NOLOCK) WHERE PARC_NumDocumento LIKE '" & vCPF & "%' AND FPAG_ID = 3 AND Len(PARC_NumDocumento) = 13 UNION SELECT BOAV_NumDocumento NumDocumento FROM Boletos_Avulsos WITH (NOLOCK) WHERE BOAV_NumDocumento LIKE '" & vCPF & "%' AND Len(BOAV_NumDocumento) = 13 UNION SELECT BOGE_Numero NumDocumento FROM Boletos_Gerados WITH (NOLOCK) WHERE BOGE_Numero LIKE '" & vCPF & "%' AND Len(BOGE_Numero) = 13) res")
			if IsNull(rsTemp("NumDocumento")) then
				vNumDoc = vCPF & "001"
			else
				vNumDoc = rsTemp("NumDocumento")
			end if
		else
			Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT PARC_NumDocumento NumDocumento FROM Parcelas WITH (NOLOCK) WHERE PARC_NumDocumento LIKE '" & vCPF & "%' AND FPAG_ID = 3 AND Len(PARC_NumDocumento) = 8 UNION SELECT BOAV_NumDocumento NumDocumento FROM Boletos_Avulsos WITH (NOLOCK) WHERE BOAV_NumDocumento LIKE '" & vCPF & "%' AND Len(BOAV_NumDocumento) = 8 UNION SELECT BOGE_Numero NumDocumento FROM Boletos_Gerados WITH (NOLOCK) WHERE BOGE_Numero LIKE '" & vCPF & "%' AND Len(BOGE_Numero) = 8) res")
			'Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT PARC_NumDocumento NumDocumento FROM Parcelas WITH (NOLOCK) WHERE PARC_NumDocumento LIKE '" & vCPF & "%' AND FPAG_ID = 3 AND Len(PARC_NumDocumento) = 8 UNION SELECT BOAV_NumDocumento NumDocumento FROM Boletos_Avulsos WITH (NOLOCK) WHERE BOAV_NumDocumento LIKE '" & vCPF & "%' AND Len(BOAV_NumDocumento) = 8) res")
			'Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT PARC_NumDocumento NumDocumento FROM Parcelas WITH (NOLOCK) WHERE PARC_NumDocumento LIKE '" & Session("FILI_ID") & "%' AND FPAG_ID = 3 AND Len(PARC_NumDocumento) >= 8 AND Len(PARC_NumDocumento) <=9 UNION SELECT BOAV_NumDocumento NumDocumento FROM Boletos_Avulsos WITH (NOLOCK) WHERE BOAV_NumDocumento LIKE '" & Session("FILI_ID") & "%' AND Len(BOAV_NumDocumento) >= 8 AND Len(BOAV_NumDocumento) <=9) res")
			'Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT PARC_NumDocumento NumDocumento FROM Parcelas WITH (NOLOCK) WHERE PARC_NumDocumento LIKE '" & Session("FILI_ID") & "%' AND FPAG_ID = 3 AND Len(PARC_NumDocumento) >= 7 UNION SELECT BOAV_NumDocumento NumDocumento FROM Boletos_Avulsos WITH (NOLOCK) WHERE BOAV_NumDocumento LIKE '" & Session("FILI_ID") & "%' AND Len(BOAV_NumDocumento) >= 7) res")
			if IsNull(rsTemp("NumDocumento")) then
				vNumDoc = vCPF & "001"
			else
				vNumDoc = rsTemp("NumDocumento")
			end if
		end if

		vID = RetornaProximoNumero("Boletos_Avulsos", "BOAV_ID", Session("FILI_ID"))
		BD.Execute "INSERT INTO Boletos_Avulsos Values(" & vID & ", " & rsParcelas("ctra_id") & ", " & vE & ", '" & vNumDoc & "', 'T', " & Replace(Replace(Request.Form("txtValor" & rsParcelas("TRAN_ID")),".",""),",",".") & ", '" & Request.Form("txtVencimento" & rsParcelas("TRAN_ID")) & "', 1, '" & Request.Form("txtVencimento" & rsParcelas("TRAN_ID")) & "', " & vValTx & ", NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, " & rsParcelas("TRAN_ID") & ")"
		PreencheLOG "INSERT INTO Boletos_Avulsos Values(" & vID & ", " & rsParcelas("ctra_id") & ", " & vE & ", '" & vNumDoc & "', 'T', " & Replace(Replace(Request.Form("txtValor" & rsParcelas("TRAN_ID")),".",""),",",".") & ", '" & Request.Form("txtVencimento" & rsParcelas("TRAN_ID")) & "', 1, '" & Request.Form("txtVencimento" & rsParcelas("TRAN_ID")) & "', " & vValTx & ", NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, " & rsParcelas("TRAN_ID") & ")", vFilialDestino

		BD.Execute "INSERT INTO Boletos_Gerados(CTRA_ID, BOGE_Numero, BOGE_TipoBoleto) VALUES(" & rsParcelas("CTRA_ID") & ", " & vNumDoc & ", 'T')"
		PreencheLOG "INSERT INTO Boletos_Gerados(CTRA_ID, BOGE_Numero, BOGE_TipoBoleto) VALUES(" & rsParcelas("CTRA_ID") & ", " & vNumDoc & ", 'T')", vFilialDestino

		vBoletos = vBoletos & vID & ","

		rsParcelas.MoveNext
	Loop
	vBoletos = left(vBoletos, len(vBoletos) - 1)
	Set rsParcelas = BD.Execute("SELECT 0 PARC_ID, b.CTRA_ID, d.DEVE_ID, DEVE_Nome, DEVE_CGCCPF, BOAV_NumDocumento PARC_NumDocumento, EDEV_ID, CTRA_Numero, b.BOAV_QtdParcelas Plano, BOAV_TaxaBoleto, BOAV_Validade, BOAV_Validade, BOAV_Vencimento, BOAV_Valor, CART_ID, CTRA_VencDebito, SCON_ID, b.TRAN_ID, TRAN_NumTitulo, CTRA_Conta FROM Boletos_Avulsos b WITH (NOLOCK) JOIN Contratos c WITH (NOLOCK) ON b.CTRA_ID = c.CTRA_ID JOIN Devedores d WITH (NOLOCK) ON c.DEVE_ID = d.DEVE_ID JOIN Transacoes t WITH (NOLOCK) ON b.TRAN_ID = t.TRAN_ID WHERE BOAV_ID IN (" & vBoletos & ") AND ACMD_ID IS NULL")
	
elseif vTipoBoleto = 2 then
	
	vTpBoleto = "A"

	Set rsCPF = BD.Execute("SELECT DEVE_CGCCPF, CTRA_Numero FROM Devedores d WITH (NOLOCK) JOIN Contratos c WITH (NOLOCK) ON c.DEVE_ID = d.DEVE_ID WHERE CTRA_ID = " & vCTRA_ID)
	if vCONT_ID = 74 or vCONT_ID = 75 then
		if len(rsCPF("DEVE_CGCCPF")) = 14 then
			vCPF = "0" & Mid(rsCPF("DEVE_CGCCPF"), 1, 8)
		else
			vCPF = Mid(rsCPF("DEVE_CGCCPF"), 1, 9)
		end if
	elseif vCONT_ID = 69 then
		vCPF = CDbl(rsCPF("DEVE_CGCCPF"))
		vCPF = Mid(vCPF & "0000", 1, 4)
	elseif vCONT_ID = 77 then
		vCPF = rsCPF("CTRA_Numero")
		vCPF = Right("00000000000000" & vCPF, 14)
	else
		vCPF = CDbl(rsCPF("DEVE_CGCCPF"))
		vCPF = Mid(vCPF & "00000", 1, 5)
	end if

	'Laser Avulso
	Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT PARC_NumDocumento NumDocumento FROM Parcelas WITH (NOLOCK) WHERE PARC_NumDocumento LIKE '" & vCPF & "%' AND FPAG_ID = 3 AND Len(PARC_NumDocumento) = 8 UNION SELECT BOAV_NumDocumento NumDocumento FROM Boletos_Avulsos WITH (NOLOCK) WHERE BOAV_NumDocumento LIKE '" & vCPF & "%' AND Len(BOAV_NumDocumento) = 8 UNION SELECT BOGE_Numero NumDocumento FROM Boletos_Gerados WITH (NOLOCK) WHERE BOGE_Numero LIKE '" & vCPF & "%' AND Len(BOGE_Numero) = 8) res")
	Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT PARC_NumDocumento NumDocumento FROM Parcelas WITH (NOLOCK) WHERE PARC_NumDocumento LIKE '" & vCPF & "%' AND FPAG_ID = 3 AND Len(PARC_NumDocumento) = 8 UNION SELECT BOAV_NumDocumento NumDocumento FROM Boletos_Avulsos WITH (NOLOCK) WHERE BOAV_NumDocumento LIKE '" & vCPF & "%' AND Len(BOAV_NumDocumento) = 8) res")
	'Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT PARC_NumDocumento NumDocumento FROM Parcelas WITH (NOLOCK) WHERE PARC_NumDocumento LIKE '" & Session("FILI_ID") & "%' AND FPAG_ID = 3 AND Len(PARC_NumDocumento) >= 8 AND Len(PARC_NumDocumento) <= 9 UNION SELECT BOAV_NumDocumento NumDocumento FROM Boletos_Avulsos WITH (NOLOCK) WHERE BOAV_NumDocumento LIKE '" & Session("FILI_ID") & "%' AND Len(BOAV_NumDocumento) >= 8 AND Len(BOAV_NumDocumento) <= 9) res")
	'Set rsTemp = BD.Execute("SELECT MAX(NumDocumento) + 1 NumDocumento FROM (SELECT PARC_NumDocumento NumDocumento FROM Parcelas WITH (NOLOCK) WHERE PARC_NumDocumento LIKE '" & Session("FILI_ID") & "%' UNION SELECT BOAV_NumDocumento NumDocumento FROM Boletos_Avulsos WITH (NOLOCK) WHERE BOAV_NumDocumento LIKE '" & Session("FILI_ID") & "%') res")
	if IsNull(rsTemp("NumDocumento")) then
		vNumDoc = vCPF & "001"
	else
		vNumDoc = rsTemp("NumDocumento")
	end if
	vID = RetornaProximoNumero("Boletos_Avulsos", "BOAV_ID", Session("FILI_ID"))
	BD.Execute "INSERT INTO Boletos_Avulsos Values(" & vID & ", " & Request.Form("ctra_id") & ", '" & vNumDoc & "', 'A', " & Replace(Replace(CDbl(vValorTaxaBoleto) + CDbl(Request.Form("txtValor")),".",""),",",".") & ", '" & Request.Form("txtVencimento") & "', " & Request.Form("txtPlano") & ", '" & DateAdd("d", Request.Form("txtValidade"), Request.Form("txtVencimento")) & "', NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL)"
	PreencheLOG "INSERT INTO Boletos_Avulsos Values(" & vID & ", " & Request.Form("ctra_id") & ", '" & vNumDoc & "', 'A', " & Replace(Replace(CDbl(vValorTaxaBoleto) + CDbl(Request.Form("txtValor")),".",""),",",".") & ", '" & Request.Form("txtVencimento") & "', " & Request.Form("txtPlano") & ", '" & DateAdd("d", Request.Form("txtValidade"), Request.Form("txtVencimento")) & "', NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL, NULL)", vFilialDestino

	BD.Execute "INSERT INTO Boletos_Gerados(CTRA_ID, BOGE_Numero, BOGE_TipoBoleto) VALUES(" & rsParcelas("CTRA_ID") & ", " & vNumDoc & ", 'A')"
	PreencheLOG "INSERT INTO Boletos_Gerados(CTRA_ID, BOGE_Numero, BOGE_TipoBoleto) VALUES(" & rsParcelas("CTRA_ID") & ", " & vNumDoc & ", 'A')", vFilialDestino

	Set rsParcelas = BD.Execute("SELECT d.DEVE_ID, DEVE_Nome, " & vNumDoc & " PARC_NumDocumento, " & vE & " EDEV_ID, CART_ID, SCON_ID, CTRA_Conta FROM Contratos c WITH (NOLOCK) JOIN Devedores d WITH (NOLOCK) ON c.DEVE_ID = d.DEVE_ID WHERE CTRA_ID = " & Request.Form("ctra_id"))
end if

Dim vHora, vValid, vPrazoPag, vUso_Banco
Dim rsPol, vJurosAtraso
Dim rsOutrosContratos, vContratos
Dim vQtdContratosHSBC
vHora = Now
vJurosAtraso = false
Dim rsEmpresa, rsTaxaBoleto, vNumCarteira
Dim vContIdentArq
Dim vNomeEmpresa
Do While Not rsParcelas.EOF
	Set rsFilial = BD.Execute("SELECT FILI_ID, CONT_ID, CART_Numero, CART_Descricao FROM Carteiras ca WITH (NOLOCK) JOIN Contratos c WITH (NOLOCK) ON c.CART_ID = ca.CART_ID WHERE CTRA_ID = " & rsParcelas("CTRA_ID"))
	vFilialDestino = rsFilial("FILI_ID")
	
	vCont_id = rsFilial("CONT_ID")
	vContratante = rsFilial("CONT_ID")
	
	vNumCarteira = rsFilial("CART_Numero")
	Dim vDescCert
	vDescCert = rsFilial("CART_Descricao")
	
	rsFilial.Close
	Set rsFilial = Nothing

	if Request.QueryString("automatico") <> "" then
		Set rsBanco = BD.Execute("SELECT BANC_ID BANC_ID_Boleto, COBO_Agencia CONT_AgenciaBoleto, COBO_Contacorrente CONT_ContaCorrenteBoleto, COBO_CodigoCliente CONT_CodigoCliente, COBO_Carteira CONT_CarteiraBoleto, COBO_OperacaoBoleto CONT_OperacaoBoleto FROM Configuracoes_de_Boleto WITH (NOLOCK) WHERE COBO_ID = " & rsParcelas("COBO_ID"))
		Set rsEmpresa = BD.Execute("SELECT CONT_Fantasia, CONT_IdentArq FROM Contratante WITH (NOLOCK) WHERE CONT_ID = " & vContratante)
		Set rsTaxaBoleto = BD.Execute("SELECT PNEG.* FROM Politica_de_Negociacao PNEG WITH (NOLOCK) JOIN REL_PolNeg_Cart REL WITH (NOLOCK) ON REL.PNEG_ID = PNEG.PNEG_ID WHERE CART_ID = " & rsParcelas("CART_ID"))
 
		vContIdentArq = rsEmpresa("CONT_IdentArq")
		vNomeEmpresa = rsEmpresa("CONT_Fantasia")
		
		if not rsTaxaBoleto.eof then
			
			if  rsTaxaBoleto("PNEG_TaxaBoleto") then
				vValorTaxaBoleto = rsTaxaBoleto("PNEG_ValorTaxaBoleto")
			else
				vValorTaxaBoleto = 0
			end if
		
		else

				vValorTaxaBoleto = 0
				
		end if		
		
	elseif Request.QueryString("atraso") <> "" then
		Set rsBanco = BD.Execute("SELECT BANC_ID_Boleto, CONT_AgenciaBoleto, CONT_ContaCorrenteBoleto, CONT_CodigoCliente, CONT_CarteiraBoleto, CONT_Fantasia, CONT_OperacaoBoleto FROM Contratante WITH (NOLOCK) WHERE CONT_ID = " & vContratante)
		Set rsEmpresa = BD.Execute("SELECT CONT_Fantasia, CONT_IdentArq FROM Contratante WITH (NOLOCK) WHERE CONT_ID = " & vContratante)
		vContIdentArq = rsEmpresa("CONT_IdentArq")
		vNomeEmpresa = rsEmpresa("CONT_Fantasia")
	elseif Request.QueryString("recuperador") <> "" then
		Set rsBanco = BD.Execute("SELECT BANC_ID_Boleto, CONT_AgenciaBoleto, CONT_ContaCorrenteBoleto, CONT_CodigoCliente, CONT_CarteiraBoleto, CONT_Fantasia, CONT_OperacaoBoleto FROM Contratante WITH (NOLOCK) WHERE CONT_ID = " & vContratante)
		Set rsEmpresa = BD.Execute("SELECT CONT_Fantasia, CONT_IdentArq FROM Contratante WITH (NOLOCK) WHERE CONT_ID = " & vContratante)
		Set rsTaxaBoleto = BD.Execute("SELECT PNEG.* FROM Politica_de_Negociacao PNEG WITH (NOLOCK) JOIN REL_PolNeg_Cart REL WITH (NOLOCK) ON REL.PNEG_ID = PNEG.PNEG_ID WHERE CART_ID = " & rsParcelas("CART_ID"))
		vContIdentArq = rsEmpresa("CONT_IdentArq")
		vNomeEmpresa = rsEmpresa("CONT_Fantasia")
		if rsTaxaBoleto("PNEG_TaxaBoleto") then
			vValorTaxaBoleto = rsTaxaBoleto("PNEG_ValorTaxaBoleto")
		else
			vValorTaxaBoleto = 0
		end if
	else
		Set rsBanco = BD.Execute("SELECT BANC_ID_Boleto, CONT_AgenciaBoleto, CONT_ContaCorrenteBoleto, CONT_CodigoCliente, CONT_CarteiraBoleto, CONT_Fantasia, CONT_OperacaoBoleto, CONT_IdentArq FROM Contratante WITH (NOLOCK) WHERE CONT_ID = " & vContratante)
		vContIdentArq = rsBanco("CONT_IdentArq")
		vNomeEmpresa = rsBanco("CONT_Fantasia")
	end if
	
	'********************************
	' CONSTANTES
	'********************************
	
	if Session("CodigoCliente") = 17 and vCont_id = 11 then
		if vFilialDestino = 1 then
			cons_agencia = "4130"
			cons_conta = "0020040"
			cons_codcliente = ""
			cons_carteira = "06"
		elseif vFilialDestino = 2 then
			cons_agencia = "4130"
			cons_conta = "0020036"
			cons_codcliente = ""
			cons_carteira = "06"
		elseif vFilialDestino = 3 then
			cons_agencia = "4130"
			cons_conta = "0020041"
			cons_codcliente = ""
			cons_carteira = "06"
		elseif vFilialDestino = 4 then
			cons_agencia = "4130"
			cons_conta = "0020038"
			cons_codcliente = ""
			cons_carteira = "06"
		elseif vFilialDestino = 5 then
			cons_agencia = "4130"
			cons_conta = "0020037"
			cons_codcliente = ""
			cons_carteira = "06"
		elseif vFilialDestino = 6 then
			cons_agencia = "4130"
			cons_conta = "0020039"
			cons_codcliente = ""
			cons_carteira = "06"
		end if
	else
		cons_agencia = Trim(rsBanco("CONT_AgenciaBoleto"))
		cons_conta = Trim(rsBanco("CONT_ContaCorrenteBoleto"))
		cons_codcliente = Trim(rsBanco("CONT_CodigoCliente"))
		cons_carteira = Trim(rsBanco("CONT_CarteiraBoleto"))
	end if
	cons_moeda = "9"
	cons_especie = "R$"
	'cons_cedente = rsEscritorio("ESCR_NomeFantasia") & " (" & rsBanco("CONT_Fantasia") & ")"
	
	if vContratante = 83 then
		cons_cedente = " BANCO GE CAPITAL S/A " 
		cons_cedente2 = rsEscritorio("ESCR_NomeFantasia")
	elseif Session("CodigoCliente") = 3 and (vContratante <> 40 and vContratante <> 5 and vContratante <> 11) then
		cons_cedente = " Rede Capta – Esc de Adv Aury Silva S/C " 
		cons_cedente2 = rsEscritorio("ESCR_NomeFantasia")
	elseif Session("CodigoCliente") = 4 and (vContratante <> 40 and vContratante <> 5 and vContratante <> 11) then
		cons_cedente = " Rede Capta – Esc de Adv Aury Silva S/C " 
		cons_cedente2 = rsEscritorio("ESCR_NomeFantasia")
	elseif Session("CodigoCliente") = 5 and (vContratante <> 40 and vContratante <> 5 and vContratante <> 11) then
		cons_cedente = " Rede Capta – Esc de Adv Aury Silva S/C " 
		cons_cedente2 = rsEscritorio("ESCR_NomeFantasia")
	elseif Session("CodigoCliente") = 6 and (vContratante <> 40 and vContratante <> 5 and vContratante <> 11) then
		cons_cedente = " Rede Capta – Esc de Adv Aury Silva S/C " 
		cons_cedente2 = rsEscritorio("ESCR_NomeFantasia")
	elseif Session("CodigoCliente") = 7 and (vContratante <> 40 and vContratante <> 5 and vContratante <> 11) then
		cons_cedente = " Rede Capta – Esc de Adv Aury Silva S/C " 
		cons_cedente2 = rsEscritorio("ESCR_NomeFantasia")
	elseif Session("CodigoCliente") = 8 and (vContratante <> 40 and vContratante <> 5 and vContratante <> 11) then
		cons_cedente = " Rede Capta – Esc de Adv Aury Silva S/C " 
		cons_cedente2 = rsEscritorio("ESCR_NomeFantasia")
	elseif Session("CodigoCliente") = 9 and (vContratante <> 40 and vContratante <> 5 and vContratante <> 11) then
		cons_cedente = " Rede Capta – Esc de Adv Aury Silva S/C " 
		cons_cedente2 = rsEscritorio("ESCR_NomeFantasia")
	elseif Session("CodigoCliente") = 10 and (vContratante <> 40 and vContratante <> 5 and vContratante <> 11) then
		cons_cedente = " Rede Capta – Esc de Adv Aury Silva S/C " 
		cons_cedente2 = rsEscritorio("ESCR_NomeFantasia")
	elseif Session("CodigoCliente") = 11 and (vContratante <> 40 and vContratante <> 5 and vContratante <> 11) then
		cons_cedente = " Rede Capta – Esc de Adv Aury Silva S/C " 
		cons_cedente2 = rsEscritorio("ESCR_NomeFantasia")
	elseif Session("CodigoCliente") = 15 and (vContratante <> 40 and vContratante <> 5 and vContratante <> 11) then
		cons_cedente = " Rede Capta – Esc de Adv Aury Silva S/C " 
		cons_cedente2 = rsEscritorio("ESCR_NomeFantasia")
	elseif Session("CodigoCliente") = 22 and (vContratante = 3 or vContratante = 4) then
		cons_cedente = " HSBC Central de Cobrança - Banco e Cartão " 
		cons_cedente2 = rsEscritorio("ESCR_NomeFantasia")
	else
		cons_cedente = rsEscritorio("ESCR_RazaoSocial") 
		cons_cedente2 = rsEscritorio("ESCR_NomeFantasia")
	end if
	
	cons_operacao =  Trim(rsBanco("CONT_OperacaoBoleto"))
	cons_dadoscedente = "" 

	vQtdboletos = vQtdboletos + 1
	Set rsEnderecoDevedor = BD.Execute("SELECT * FROM Endereco_Devedor e WITH (NOLOCK) WHERE EDEV_ID = " & rsParcelas("EDEV_ID"))
	
	'********************************
	' VARIÁVEIS 
	'********************************
	if vCONT_ID = 58 then
		vParcelasNeg = ", parcela(s) "
		Set rsTransac = BD.Execute("SELECT TRAN_NumTitulo FROM Transacoes WITH (NOLOCK) WHERE CTRA_ID = " & rsParcelas("CTRA_ID"))
		Do While Not rsTransac.EOF
			vParcelasNeg = vParcelasNeg & rsTransac("TRAN_NumTitulo") & ", "
			rsTransac.MoveNext
		Loop

		vParcelasNeg = Left(vParcelasNeg, len(vParcelasNeg) - 2)
	end if
	if vCONT_ID = 58 or vCONT_ID = 74 or vCONT_ID = 75 then
		if len(rsParcelas("DEVE_CGCCPF")) = 11 then
			var_sacado = rsParcelas("DEVE_Nome") & " - CPF: " & FormataCPFCNPJ(rsParcelas("DEVE_CGCCPF")) '"Luiz Daniel de Souza"
		else
			var_sacado = rsParcelas("DEVE_Nome") & " - CNPJ: " & FormataCPFCNPJ(rsParcelas("DEVE_CGCCPF")) '"Luiz Daniel de Souza"
		end if
	else
		var_sacado = rsParcelas("DEVE_Nome") '"Luiz Daniel de Souza"
	end if
	var_endereco = rsEnderecoDevedor("EDEV_Endereco")  '"Avenida Data Tech Tecnologia e Informática, 2002"
	if Not IsNull(rsEnderecoDevedor("EDEV_Numero")) then
		var_endereco = var_endereco & ", " & rsEnderecoDevedor("EDEV_Numero")
	end if
	if Not IsNull(rsEnderecoDevedor("EDEV_Complemento")) then
		var_endereco = var_endereco & ", " & rsEnderecoDevedor("EDEV_Complemento")
	end if
	var_bairro = rsEnderecoDevedor("EDEV_Bairro") '"Serpa"
	var_cidade = rsEnderecoDevedor("EDEV_Cidade") '"Caieiras"
	var_estado = rsEnderecoDevedor("ESTA_ID") '"São Paulo"
	var_cep = FormataCEP(rsEnderecoDevedor("EDEV_CEP")) '"07700-000"

	var_datadocumento = date() '"03/06/2002"

	vUso_Banco = ""
	if rsBanco("BANC_ID_Boleto") = "008" or rsBanco("BANC_ID_Boleto") = "353" then
		var_nossonumero = Right("000000000000" & rsParcelas("PARC_NumDocumento"), 12)
		var_nomebanco = "Santander"
		'Código do Cedente com 7 dígitos
	elseif rsBanco("BANC_ID_Boleto") = "033" then
		var_nossonumero = Right("0000000" & rsParcelas("PARC_NumDocumento"), 7)
		var_nomebanco = "Banespa"
	elseif rsBanco("BANC_ID_Boleto") = 104 then
		var_nossonumero = Right("00000000000000" & rsParcelas("PARC_NumDocumento"), 14)
		var_nomebanco = "Caixa"
		'Agência com 4 dígitos
		'Código do Cedente com 5 dígitos
	elseif rsBanco("BANC_ID_Boleto") = 237 then
		var_nossonumero = Right("00000000000" & rsParcelas("PARC_NumDocumento"), 11)
		var_nomebanco = "Bradesco"
		'Agência com 4 dígitos
		'Conta corrente com 7 dígitos
	elseif rsBanco("BANC_ID_Boleto") = 341 then
		var_nossonumero = Right("00000000" & rsParcelas("PARC_NumDocumento"), 8)
		var_nomebanco = "Itaú"
		'Agência com 4 dígitos
		'Conta corrente com 5 dígitos
	elseif rsBanco("BANC_ID_Boleto") = 347 then
		var_nossonumero = Right("0000000000000" & rsParcelas("PARC_NumDocumento"), 13)
		var_nomebanco = "Banco Sudameris"
		'Agência com 4 dígitos
		'Conta corrente com 7 dígitos
	elseif rsBanco("BANC_ID_Boleto") = 356 then
		var_nossonumero = Right("0000000000000" & rsParcelas("PARC_NumDocumento"), 13)
		var_nomebanco = "Banco Real"
		'Agência com 4 dígitos
		'Conta corrente com 7 dígitos
	elseif rsBanco("BANC_ID_Boleto") = 399 then
		var_nossonumero = Right("0000000000000" & rsParcelas("PARC_NumDocumento"), 13)
		var_nomebanco = "HSBC"
	elseif rsBanco("BANC_ID_Boleto") = 409 then
		var_nossonumero = Right("00000000000000" & rsParcelas("PARC_NumDocumento"), 14)
		var_nomebanco = "Unibanco"
		vUso_Banco = "CVT 77445"
		'Agência com 4 dígitos
		'Conta corrente com 8 dígitos
		'Código do cliente com 7 dígitos
	elseif rsBanco("BANC_ID_Boleto") = 623 then
		var_nossonumero = Right("0000000000" & rsParcelas("PARC_NumDocumento"), 10)
		var_nomebanco = "PANAMERICANO"
		'Agência com 4 dígitos
		'Conta corrente com 5 dígitos
		'Código do cliente com 7 dígitos
	end if


	if Session("CodigoCliente") = 23 and vNumCarteira = "3" then
		vFilialDestino = 3
	end if
	Set rsFilial2 = BD.Execute("SELECT * FROM Filiais WITH (NOLOCK) WHERE FILI_ID = " & vFilialDestino)
	vTel = rsFilial2("FILI_Tel1")
	vFax = rsFilial2("FILI_Fax")
	vEndFilial = rsFilial2("FILI_Endereco") & " - " & rsFilial2("FILI_Bairro") & " - " & rsFilial2("FILI_Cidade") & " - " & rsFilial2("FILI_Estado") & " CEP: " & rsFilial2("FILI_CEP")
 
  
	
	Set rsPol = BD.Execute("SELECT PNEG_CorrigeParcelaAtraso, PNEG_JurosParcela FROM Politica_de_Negociacao p WITH (NOLOCK) JOIN REL_PolNeg_Cart r WITH (NOLOCK) ON p.PNEG_ID = r.PNEG_ID WHERE CART_ID = " & rsParcelas("CART_ID"))
	
	
	
	vHora = DateAdd("s", 1, vHora)
	if Request.QueryString("recuperador") <> "" then
		if vCont_id = 43 then
			var_contrato = rsParcelas("CTRA_Conta") & rsParcelas("CTRA_Numero") & " - " & rsEmpresa("CONT_Fantasia")
		elseif rsParcelas("CART_ID") = 85 then
			Dim rsNumTituloCt4, strDescTituloCt4
			Set rsNumTituloCt4 = Bd.Execute("SELECT * FROM Transacoes_de_Acordo t WITH (NOLOCK) LEFT JOIN Tipos_de_Documento td WITH (NOLOCK) ON t.TDOC_ID = td.TDOC_ID WHERE ACOR_ID = " & rsParcelas("ACOR_ID"))
			if rsNumTituloCt4.EOF then
				Set rsNumTituloCt4 = Bd.Execute("SELECT * FROM Transacoes t WITH (NOLOCK) LEFT JOIN Tipos_de_Documento td WITH (NOLOCK) ON t.TDOC_ID = td.TDOC_ID WHERE CTRA_ID = " & rsParcelas("CTRA_ID"))
			end if
			strDescTituloCt4 = ""
			Do while not rsNumTituloCt4.EOF
				strDescTituloCt4 = strDescTituloCt4 & rsNumTituloCt4("TRAN_NumTitulo") & ", "	
				rsNumTituloCt4.movenext
			loop
			strDescTituloCt4 = Left(strDescTituloCt4, len(strDescTituloCt4) - 2)
			var_contrato = strDescTituloCt4			
		else
			var_contrato = FormataContrato(rsParcelas("CTRA_Numero")) & " - " & rsEmpresa("CONT_Fantasia")
		end if
		var_parcelaplano = "1 / " & rsParcelas("Plano")

		var_datavencimento = rsParcelas("BOAV_Vencimento") '"05/06/2002"
		var_valordocumento = FormatNumber(vValorTaxaBoleto + rsParcelas("BOAV_Valor")) '"70,00"

		if Session("CodigoCliente") = 13 and vCONT_ID = 11 then
			vTel = "(21)3212-3771"
		elseif Session("CodigoCliente") = 13 and (rsParcelas("CART_ID") = 38001 or rsParcelas("CART_ID") = 47001 or rsParcelas("CART_ID") = 42001 or rsParcelas("CART_ID") = 48001) then
			vTel = "(21)3212-3792"
		elseif Session("CodigoCliente") = 13 and vCONT_ID = 27 then
			vTel = "(21)3212-3780"
		elseif Session("CodigoCliente") = 13 and vCONT_ID = 14 then
			vTel = "(21)3212-3788"
		elseif Session("CodigoCliente") = 13 and (vCONT_ID = 16 or vCONT_ID = 17 or vCONT_ID = 61) then
			vTel = "(21)3212-3789"
		elseif Session("CodigoCliente") = 13 and vCONT_ID = 77 then
			vTel = "(21)3212-3793"
		elseif Session("CodigoCliente") = 13 and vCONT_ID = 80 then
			vTel = "(21)3212-3794"
		elseif Session("CodigoCliente") = 13 and vCONT_ID = 86 then
			vTel = "(21)3212-3796"
		elseif Session("CodigoCliente") = 13 and vCONT_ID = 48 then
			vTel = "(21)3212-3786"
		elseif Session("CodigoCliente") = 1 and vCONT_ID = 85 then
			vTel = "(71)3483-1250"
		elseif Session("CodigoCliente") = 23 and vCONT_ID = 58 then
			vTel = "(21)3984-2705"
		elseif Session("CodigoCliente") = 25 and vCONT_ID = 54 then
			vTel = "São Paulo (11) 2165-9423 - Demais localidades 0800 724 23 23"
		elseif Session("CodigoCliente") = 25 and (vCONT_ID = 100 or vCONT_ID = 104) then
			vTel = "SP e Gde. SP: (11) 2165-9436 - Outras Localidades:  0800-7241100"
		end if
		if vTel <> "" then
			vTel = " Tel.: " & vTel
		end if
		if vFax <> "" then
			vFax = " Fax.: " & vFax
		end if

		vValidadeBoleto = rsParcelas("BOAV_Validade")
		if vCont_id = 1 then
			if CDbl(rsParcelas("BOAV_TaxaBoleto")) > 0 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & rsParcelas("BOAV_Validade") & ". <BR>Após " & rsParcelas("BOAV_Validade") & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & ")<BR>Valor do documento acrescido de " & FormatCurrency(rsParcelas("BOAV_TaxaBoleto")) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " <br>O não pagamento até a data de vencimento acarretará no cancelamento do acordo e a inclusão do seu nome no órgão proteção ao crédito de sua região e no SERASA. Implicando na perda do desconto porventura concedido.</b>" 
			else
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & rsParcelas("BOAV_Validade") & ". <BR>Após " & rsParcelas("BOAV_Validade") & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & ")<BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " <br>O não pagamento até a data de vencimento acarretará no cancelamento do acordo e a inclusão do seu nome no órgão proteção ao crédito de sua região e no SERASA. Implicando na perda do desconto porventura concedido.</b>" 
			end if
		elseif vCont_id = 8 OR vCont_id = 9 then
			if CDbl(rsParcelas("BOAV_TaxaBoleto")) > 0 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & rsParcelas("BOAV_Validade") & ". <BR>Após " & rsParcelas("BOAV_Validade") & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & ")<BR>Valor do documento acrescido de " & FormatCurrency(rsParcelas("BOAV_TaxaBoleto")) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " <br>O não pagamento até a data de vencimento acarretará no cancelamento do acordo e a inclusão do seu nome no órgão proteção ao crédito de sua região.</b>" 
			else
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & rsParcelas("BOAV_Validade") & ". <BR>Após " & rsParcelas("BOAV_Validade") & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & ")<BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " <br>O não pagamento até a data de vencimento acarretará no cancelamento do acordo e a inclusão do seu nome no órgão proteção ao crédito de sua região.</b>" 
			end if
		elseif vCont_id = 2 then
			var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & rsParcelas("BOAV_Validade") & ". <BR>Após o vencimento cobrar 2% de multa. <BR>Após " & rsParcelas("BOAV_Validade") & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " </b>" 
		elseif vCont_id = 43 then
			 var_instrucoes = "<B>Pagável em qualquer banco até o dia " & rsParcelas("BOAV_Validade") & ". <BR>Após " & rsParcelas("BOAV_Validade") & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>CGC/CPF: " & FormataCPFCNPJ(rsParcelas("DEVE_CGCCPF")) & " - Parcela/Plano: " & var_parcelaplano & ".<br></b>&nbsp; Titulo(s):&nbsp;" & rsParcelas("BOAV_Texto2") & ". Sr. Cliente Caso tenha ocorrido mudança de endereço, favor comparecer agência para atualização dos dados cadastrais para que o carnê com os próximos pagamentos chegue no endereço correto." 
		elseif vCont_id = 3 then
			vQtdContratosHSBC = 0
			vContratos = rsParcelas("BOAV_Texto2")
			vContratos = vContratos & " - " & rsBanco("CONT_Fantasia")
			var_contrato = "&nbsp;"
			var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & rsParcelas("BOAV_Validade") & ". <BR>Após " & rsParcelas("BOAV_Validade") & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Referente ao(s) contrato(s) nº: " & vContratos & "<br>Parcela/Plano: " & var_parcelaplano & " - CNPJ/CPF: " & FormataCPFCNPJ(rsParcelas("DEVE_CGCCPF")) & "<br>O envio desse boleto não inibe uma possível Ação de Débito efetuado pelo HSBC na conta corrente.</b>" 
		elseif vCont_id = 4 then
			var_contrato = Mid(var_contrato,1 ,2) & "**.****.****." & Mid(var_contrato,16 ,4)
			var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & rsParcelas("BOAV_Validade") & ". <BR>Após " & rsParcelas("BOAV_Validade") & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Referente ao contrato nº: " & var_contrato & "<br>Parcela/Plano: " & var_parcelaplano & " - CNPJ/CPF: " & FormataCPFCNPJ(rsParcelas("DEVE_CGCCPF")) & " <br>O envio desse boleto não inibe uma possível Ação de Débito efetuado pelo HSBC na conta corrente.</b>" 
		elseif vCont_id = 11 and rsParcelas("Plano") = 1 then
			if CDbl(rsParcelas("BOAV_TaxaBoleto")) > 0 then
				var_instrucoes = "<BR> <B>Pagável apenas até o vencimento.<BR>Após o vencimento, entrar em contato com a " & cons_cedente2 & " - " & vTel & " " & vFax & ".<BR> Para pagamento em cheque, a quitação deste título ocorrerá após a quitação. <BR>Valor do documento acrescido de " & FormatCurrency(rsParcelas("BOAV_TaxaBoleto")) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & "</b>" 
			else
				var_instrucoes = "<BR> <B>Pagável apenas até o vencimento. <BR>Após o vencimento, entrar em contato com a " & cons_cedente2 & " - " & vTel & " " & vFax & ".<BR> Para pagamento em cheque, a quitação deste título ocorrerá após a quitação. <BR>Referente ao contrato nº: " & var_contrato & "</b>" 
			end if
		elseif vCont_id = 11 and rsParcelas("Plano") > 1 then
			if CDbl(rsParcelas("BOAV_TaxaBoleto")) > 0 then
				var_instrucoes = "<BR> <B>Após o vencimento, pagável somente no Banco Bradesco até o 10º dia de vencido. <br>Em caso de atraso, atualizar o valor desde a data de vencimento aplicando TR + 2,00% AM pro rata. <br>Após o 10º dia de vencido, entrar em contato com a " & cons_cedente2 & " - " & vTel & " " & vFax & ". <br> Para pagamento em cheque, a quitação deste título ocorrerá após a compensação.<BR>Valor do documento acrescido de " & FormatCurrency(rsParcelas("BOAV_TaxaBoleto")) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " </b>" 
			else
				var_instrucoes = "<BR> <B>Após o vencimento, pagável somente no Banco Bradesco até o 10º dia de vencido. <br>Em caso de atraso, atualizar o valor desde a data de vencimento aplicando TR + 2,00% AM pro rata. <br>Após o 10º dia de vencido, entrar em contato com a " & cons_cedente2 & " - " & vTel & " " & vFax & ". <br> Para pagamento em cheque, a quitação deste título ocorrerá após a compensação.<BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " </b>" 
			end if
		elseif vCont_id = 14 and Session("CodigoCliente") = 17 then
			var_instrucoes = "<BR> <B>Não receber após o vencimento.<br>Não receber valor inferior ao do documento.<br>Após " & rsParcelas("BOAV_Validade") & " entrar em contato com " & cons_cedente2 & " - " & vTel & " " & vFax & ". <br>Empresa mandatária: Cobresp.<BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " </b>" 
		elseif vCont_id = 14 then
			if CDbl(rsParcelas("BOAV_TaxaBoleto")) > 0 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & rsParcelas("BOAV_Validade") & ". <BR>Após " & rsParcelas("BOAV_Validade") & " entrar em contato com " & cons_cedente2 & " - " & vTel & " " & vFax & ". <BR>Valor do documento acrescido de " & FormatCurrency(rsParcelas("BOAV_TaxaBoleto")) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " </b>" 
			else
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & rsParcelas("BOAV_Validade") & ". <BR>Após " & rsParcelas("BOAV_Validade") & " entrar em contato com " & cons_cedente2 & " - " & vTel & " " & vFax & ". <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " </b>" 
			end if
		elseif vCont_id = 48 then
			if CDbl(vValorTaxaBoleto) > 0 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & rsParcelas("BOAV_Validade") & ". <BR>Após " & rsParcelas("BOAV_Validade") & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & ").<BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " </b>. O não pagamento até a data de vencimento acarretará no cancelamento do acordo e a inclusão do seu nome no órgão de proteção ao crédito de sua região. " 

			else
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & rsParcelas("BOAV_Validade") & ". <BR>Após " & rsParcelas("BOAV_Validade") & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & ").<BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " </b>. O não pagamento até a data de vencimento acarretará no cancelamento do acordo e a inclusão do seu nome no órgão de proteção ao crédito de sua região. " 
			end if
		elseif vCont_id = 58 then
		'Ponto Frio
			if CDbl(vValorTaxaBoleto) > 0 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & rsParcelas("BOAV_Validade") & ". <BR>Após " & rsParcelas("BOAV_Validade") & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & ").<BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & ".<br> Referente a(s) parcela(s) " & vTitulos & "</b>. O não pagamento até a data de vencimento acarretará no cancelamento do acordo e a inclusão do seu nome no órgão de proteção ao crédito de sua região. " 
			else
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & rsParcelas("BOAV_Validade") & ". <BR>Após " & rsParcelas("BOAV_Validade") & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & ").<BR>Referente ao contrato nº: " & var_contrato & ".<br> Referente a(s) parcela(s) " & vTitulos & "</b>. O não pagamento até a data de vencimento acarretará no cancelamento do acordo e a inclusão do seu nome no órgão de proteção ao crédito de sua região. " 
			end if
		elseif vCont_id = 54 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & rsParcelas("BOAV_Validade") & ". <BR>Após " & rsParcelas("BOAV_Validade") & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " </b>" 
		elseif vCONT_id = 53 then
			if CDbl(vValorTaxaBoleto) > 0 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & ". <BR>Após " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao saldo remanescente do contrato: "& var_contrato &" - Parcela/Plano: " & var_parcelaplano &". </b>" 
			else
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & ". <BR>Após " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Referente ao saldo remanescente do contrato: "& var_contrato &" - Parcela/Plano: " & var_parcelaplano & ". </b>" 
			end if
		
		else
			if CDbl(vValorTaxaBoleto) > 0 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & rsParcelas("BOAV_Validade") & ". <BR>Após " & rsParcelas("BOAV_Validade") & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " </b>" 
			else
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & rsParcelas("BOAV_Validade") & ". <BR>Após " & rsParcelas("BOAV_Validade") & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " </b>" 
			end if
		end if
		var_textoparcela = "1 / " & rsParcelas("Plano")
		
		'Preenche o andamento da cobrança
		'if Session("CodigoCliente") <> 0 then
		BD.Execute "EXEC prc_CriaAndamentoCobranca @CTRA_ID = " & rsParcelas("ctra_id") & ", @ANCO_DataHora = '" & vHora & "', @STAC_ID = 23, @TDEV_ID = NULL, @ANCO_Descricao = 'Emissão do boleto avulso " & var_nossonumero & ", no valor de R$ " & var_valordocumento & ", com vencimento em " & var_datavencimento & ", referente à parcela/plano " & var_parcelaplano & ". Boleto solicitado pelo recuperador.', @FUNC_ID = " & Session("FUNC_ID")
		PreencheLOG "EXEC prc_CriaAndamentoCobranca @CTRA_ID = " & rsParcelas("ctra_id") & ", @ANCO_DataHora = '" & vHora & "', @STAC_ID = 23, @TDEV_ID = NULL, @ANCO_Descricao = 'Emissão do boleto avulso " & var_nossonumero & ", no valor de R$ " & var_valordocumento & ", com vencimento em " & var_datavencimento & ", referente à parcela/plano " & var_parcelaplano & ". Boleto solicitado pelo recuperador.', @FUNC_ID = " & Session("FUNC_ID"), vFilialDestino
		'end if
				
		if rsParcelas("SCON_ID") <> 2 then
			vHora = DateAdd("s", 1, vHora)
			strSQL = "EXEC prc_CriaAndamentoCobranca @CTRA_ID = " & rsParcelas("ctra_id") & ", @ANCO_DataHora = '" & vHora & "', @STAC_ID = 13, @TDEV_ID = NULL, @ANCO_Descricao = 'Contrato direcionado para a fila de cobrança Pré-Acordo. Motivo: Boleto avulso emitido.', @FUNC_ID = 0"
			BD.Execute strSQL
			PreencheLOG strSQL, vFilialDestino

			BD.Execute "prc_INS_HistoricoAgendamento_Forcado 3001, '" & now & "', " & rsParcelas("ctra_id") & ", " & Session("FILI_ID") & ", " & Session("FUNC_ID")
			PreencheLOG "prc_INS_HistoricoAgendamento_Forcado 3001, '" & now & "', " & rsParcelas("ctra_id") & ", " & Session("FILI_ID") & ", " & Session("FUNC_ID"), vFilialDestino
		end if
	elseif Request.QueryString("automatico") <> "" then
		if vCont_id = 43 then
			var_contrato = rsParcelas("CTRA_Conta") & rsParcelas("CTRA_Numero") & " - " & rsEmpresa("CONT_Fantasia")
		elseif rsParcelas("CART_ID") = 85 then
			Dim rsNumTituloCt3, strDescTituloCt3
			Set rsNumTituloCt3 = Bd.Execute("SELECT * FROM Transacoes_de_Acordo t WITH (NOLOCK) LEFT JOIN Tipos_de_Documento td WITH (NOLOCK) ON t.TDOC_ID = td.TDOC_ID WHERE ACOR_ID = " & rsParcelas("ACOR_ID"))
			if rsNumTituloCt3.EOF then
				Set rsNumTituloCt3 = Bd.Execute("SELECT * FROM Transacoes t WITH (NOLOCK) LEFT JOIN Tipos_de_Documento td WITH (NOLOCK) ON t.TDOC_ID = td.TDOC_ID WHERE CTRA_ID = " & rsParcelas("CTRA_ID"))
			end if
			strDescTituloCt3 = ""
			Do while not rsNumTituloCt3.EOF
				strDescTituloCt3 = strDescTituloCt3 & rsNumTituloCt3("TRAN_NumTitulo") & ", "	
				rsNumTituloCt3.movenext
			loop
			strDescTituloCt3 = Left(strDescTituloCt3, len(strDescTituloCt3) - 2)
			var_contrato = strDescTituloCt3			
		else
			var_contrato = FormataContrato(rsParcelas("CTRA_Numero")) & " - " & rsEmpresa("CONT_Fantasia")
		end if
		var_parcelaplano = rsParcelas("PARC_Numero") & " / " & rsParcelas("Plano")

		var_datavencimento = rsParcelas("PARC_Vencimento") '"05/06/2002"
		var_valordocumento = FormatNumber(vValorTaxaBoleto + rsParcelas("PARC_ValorTotal")) '"70,00"

		if Session("CodigoCliente") = 13 and vCONT_ID = 11 then
			vTel = "(21)3212-3771"
		elseif Session("CodigoCliente") = 13 and (rsParcelas("CART_ID") = 38001 or rsParcelas("CART_ID") = 47001 or rsParcelas("CART_ID") = 42001 or rsParcelas("CART_ID") = 48001) then
			vTel = "(21)3212-3792"
		elseif Session("CodigoCliente") = 13 and vCONT_ID = 27 then
			vTel = "(21)3212-3780"
		elseif Session("CodigoCliente") = 13 and vCONT_ID = 14 then
			vTel = "(21)3212-3788"
		elseif Session("CodigoCliente") = 13 and (vCONT_ID = 16 or vCONT_ID = 17 or vCONT_ID = 61) then
			vTel = "(21)3212-3789"
		elseif Session("CodigoCliente") = 13 and vCONT_ID = 77 then
			vTel = "(21)3212-3793"
		elseif Session("CodigoCliente") = 13 and vCONT_ID = 80 then
			vTel = "(21)3212-3794"
		elseif Session("CodigoCliente") = 13 and vCONT_ID = 86 then
			vTel = "(21)3212-3796"
		elseif Session("CodigoCliente") = 13 and vCONT_ID = 48 then
			vTel = "(21)3212-3786"
		elseif Session("CodigoCliente") = 1 and vCONT_ID = 85 then
			vTel = "(71)3483-1250"
		elseif Session("CodigoCliente") = 23 and vCONT_ID = 58 then
			vTel = "(21)3984-2705"
		elseif Session("CodigoCliente") = 25 and vCONT_ID = 54 then
			vTel = "São Paulo (11) 2165-9423 - Demais localidades 0800 724 23 23"
		elseif Session("CodigoCliente") = 25 and (vCONT_ID = 100 or vCONT_ID = 104) then
			vTel = "SP e Gde. SP: (11) 2165-9436 - Outras Localidades: 0800-7241100"			
		end if

		if vTel <> "" then
			vTel = " Tel.: " & vTel
		end if
		if vFax <> "" then
			vFax = " Fax.: " & vFax
		end if

		vValidadeBoleto = DateAdd("d", request("txtvalidade"), var_datavencimento) 

		if vCont_id = 1 then
			if CDbl(vValorTaxaBoleto) > 0 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & vValidadeBoleto & ". <BR>Após " & var_datavencimento & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "). <BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " <br>O não pagamento até a data de vencimento acarretará no cancelamento do acordo e a inclusão do seu nome no órgão proteção ao crédito de sua região e no SERASA. Implicando na perda do desconto porventura concedido.</b>" 
			else
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & vValidadeBoleto & ". <BR>Após " & var_datavencimento & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "). <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " <br>O não pagamento até a data de vencimento acarretará no cancelamento do acordo e a inclusão do seu nome no órgão proteção ao crédito de sua região e no SERASA. Implicando na perda do desconto porventura concedido.</b>" 
			end if
		elseif vCont_id = 8 OR vCont_id = 9 then
			if CDbl(vValorTaxaBoleto) > 0 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & vValidadeBoleto & ". <BR>Após " & var_datavencimento & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "). <BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " <br>O não pagamento até a data de vencimento acarretará no cancelamento do acordo e a inclusão do seu nome no órgão proteção ao crédito de sua região.</b>" 
			else
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & vValidadeBoleto & ". <BR>Após " & var_datavencimento & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "). <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " <br>O não pagamento até a data de vencimento acarretará no cancelamento do acordo e a inclusão do seu nome no órgão proteção ao crédito de sua região.</b>" 
			end if
		elseif vCont_id = 2 then
			var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & vValidadeBoleto & ". <BR>Após o vencimento cobrar 2% de multa. <BR>Após " & var_datavencimento & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>CGC/CPF: " & FormataCPFCNPJ(rsParcelas("DEVE_CGCCPF")) & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " </b>" 
		elseif vCont_id = 43 then
			var_instrucoes = "<B>Pagável em qualquer banco até o dia " & vValidadeBoleto & ". <BR>Após " & var_datavencimento & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & ".<br>CGC/CPF: " & FormataCPFCNPJ(rsParcelas("DEVE_CGCCPF")) & ".</b>" 
		elseif vCont_id = 11 and rsParcelas("Plano") = 1 then
			if CDbl(vValorTaxaBoleto) > 0 then
				var_instrucoes = "<BR> <B>Pagável apenas até o vencimento.<BR>Após o vencimento, entrar em contato com a " & cons_cedente2 & " - " & vTel & " " & vFax & ".<BR> Para pagamento em cheque, a quitação deste título ocorrerá após a quitação. <BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & "</b>" 
			else
				var_instrucoes = "<BR> <B>Pagável apenas até o vencimento. <BR>Após o vencimento, entrar em contato com a " & cons_cedente2 & " - " & vTel & " " & vFax & ".<BR> Para pagamento em cheque, a quitação deste título ocorrerá após a quitação. <BR>Referente ao contrato nº: " & var_contrato & "</b>" 
			end if
		elseif vCont_id = 11 and rsParcelas("Plano") > 1 then
			if CDbl(vValorTaxaBoleto) > 0 then
				var_instrucoes = "<BR> <B>Após o vencimento, pagável somente no Banco Bradesco até o 10º dia de vencido. <br>Em caso de atraso, atualizar o valor desde a data de vencimento aplicando TR + 2,00% AM pro rata. <br>Após o 10º dia de vencido, entrar em contato com a " & cons_cedente2 & " - " & vTel & " " & vFax & ". <br> Para pagamento em cheque, a quitação deste título ocorrerá após a compensação.<BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " </b>" 
			else
				var_instrucoes = "<BR> <B>Após o vencimento, pagável somente no Banco Bradesco até o 10º dia de vencido. <br>Em caso de atraso, atualizar o valor desde a data de vencimento aplicando TR + 2,00% AM pro rata. <br>Após o 10º dia de vencido, entrar em contato com a " & cons_cedente2 & " - " & vTel & " " & vFax & ". <br> Para pagamento em cheque, a quitação deste título ocorrerá após a compensação.<BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " </b>" 
			end if
		elseif vCont_id = 14 and Session("CodigoCliente") = 17 then
			var_instrucoes = "<BR> <B>Não receber após o vencimento.<br>Não receber valor inferior ao do documento.<br>Após " & var_datavencimento & " entrar em contato com " & cons_cedente2 & " - " & vTel & " " & vFax & ". <br>Empresa mandatária: Cobresp.<BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " </b>" 
		elseif vCont_id = 14 then
			if CDbl(vValorTaxaBoleto) > 0 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & vValidadeBoleto & ". <BR>Após " & var_datavencimento & " entrar em contato com " & cons_cedente2 & " - " & vTel & " " & vFax & ". <BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " </b>" 
			else
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & vValidadeBoleto & ". <BR>Após " & var_datavencimento & " entrar em contato com " & cons_cedente2 & " - " & vTel & " " & vFax & ". <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " </b>" 
			end if
		elseif vCont_id = 48 then
			if CDbl(vValorTaxaBoleto) > 0 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & vValidadeBoleto & ". <BR>Após " & var_datavencimento & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & ").<BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " </b>. O não pagamento até a data de vencimento acarretará no cancelamento do acordo e a inclusão do seu nome no órgão de proteção ao crédito de sua região. " 
			else
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & vValidadeBoleto & ". <BR>Após " & var_datavencimento & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & ").<BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " </b>. O não pagamento até a data de vencimento acarretará no cancelamento do acordo e a inclusão do seu nome no órgão de proteção ao crédito de sua região. " 
			end if
		elseif vCont_id = 3 then
				vContratos = ""
						Set rsOutrosContratos = BD.Execute("SELECT TRAN_NumTitulo FROM Transacoes WHERE CTRA_ID = " & rsParcelas("CTRA_ID") & " AND TDOC_ID <> 307")
						do While not rsOutrosContratos.EOF 
							vContratos = vContratos & rsOutrosContratos("TRAN_NumTitulo") & ", "
							rsOutrosContratos.MoveNext
						loop
				if vContratos <> "" then		
					vContratos = Mid(vContratos, 1, len(vContratos) - 2)
				end if
			if CDbl(vValorTaxaBoleto) > 0 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & vValidadeBoleto & ". <BR>Após " & var_datavencimento & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao(s) contrato(s) nº: " & vContratos & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " - CNPJ/CPF: " & FormataCPFCNPJ(rsParcelas("DEVE_CGCCPF")) & " <br>O envio desse boleto não inibe uma possível Ação de Débito efetuado pelo HSBC na conta corrente.</b>" 
			else
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & vValidadeBoleto & ". <BR>Após " & var_datavencimento & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Referente ao(s) contrato(s) nº: " & vContratos & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " - CNPJ/CPF: " & FormataCPFCNPJ(rsParcelas("DEVE_CGCCPF")) & " <br>O envio desse boleto não inibe uma possível Ação de Débito efetuado pelo HSBC na conta corrente.</b>" 
			end if
		elseif vCont_id = 4 then
			var_contrato = Mid(var_contrato,1 ,2) & "**.****.****." & Mid(var_contrato,16 ,4)
			if CDbl(vValorTaxaBoleto) > 0 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & vValidadeBoleto & ". <BR>Após " & var_datavencimento & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " - CNPJ/CPF: " & FormataCPFCNPJ(rsParcelas("DEVE_CGCCPF")) & " <br>O envio desse boleto não inibe uma possível Ação de Débito efetuado pelo HSBC na conta corrente.</b>" 
			else
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & vValidadeBoleto & ". <BR>Após " & var_datavencimento & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " - CNPJ/CPF: " & FormataCPFCNPJ(rsParcelas("DEVE_CGCCPF")) & " <br>O envio desse boleto não inibe uma possível Ação de Débito efetuado pelo HSBC na conta corrente.</b>" 
			end if
		elseif not rsPol.eof  then
			if rsPol("PNEG_CorrigeParcelaAtraso") then
				var_instrucoes = "<BR> <B>Senhor Caixa: Cobrar " & FormatCurrency(AtualizaValorParcelaBoleto(rsParcelas("CART_ID"), CDate(rsParcelas("PARC_Vencimento")), CDate(rsParcelas("PARC_Vencimento")) + 1, rsParcelas("PARC_ValorTotal"), rsParcelas("FNEG_ID")) - rsParcelas("PARC_ValorTotal")) & " por dia de atraso.<BR>Pagável em qualquer banco até o dia " & var_datavencimento & " acrescido dos encargos por dia de atraso. <BR>Após " & var_datavencimento & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " </b>" 
			vJurosAtraso = true
			end if
		elseif vCont_id = 54 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & vValidadeBoleto & ". <BR>Após " & var_datavencimento & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " </b>" 
		elseif vCONT_id = 53 then
			if CDbl(vValorTaxaBoleto) > 0 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & ". <BR>Após " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao saldo remanescente do contrato: "& var_contrato &" - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") &". </b>" 
			else
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & ". <BR>Após " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Referente ao saldo remanescente do contrato: "& var_contrato &" - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & ". </b>" 
			end if
		else
			if CDbl(vValorTaxaBoleto) > 0 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & vValidadeBoleto & ". <BR>Após " & var_datavencimento & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " </b>" 
			else
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & vValidadeBoleto & ". <BR>Após " & var_datavencimento & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " </b>" 
			end if
		end if
		var_textoparcela = rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano")
		
		'Preenche o andamento da cobrança
		'if Session("CodigoCliente") <> 0 then
		if Session("CodigoCliente") = 20 then
			BD.Execute "EXEC prc_CriaAndamentoCobranca @CTRA_ID = " & rsParcelas("CTRA_ID") & ", @ANCO_DataHora = '" & vHora & "', @STAC_ID = 23, @TDEV_ID = NULL, @ANCO_Descricao = 'Emissão do boleto " & rsParcelas("PARC_NumDocumento") & ", referente à parcela " & rsParcelas("PARC_Numero") & ". Venct. " & var_datavencimento & ", validade " & vValidadeBoleto & ". Valor R$ "& var_valordocumento &" . Impressão automática.', @FUNC_ID = " & Session("FUNC_ID")
			PreencheLOG "EXEC prc_CriaAndamentoCobranca @CTRA_ID = " & rsParcelas("CTRA_ID") & ", @ANCO_DataHora = '" & vHora & "', @STAC_ID = 23, @TDEV_ID = NULL, @ANCO_Descricao = 'Emissão do boleto " & rsParcelas("PARC_NumDocumento") & ", referente à parcela " & rsParcelas("PARC_Numero") & ". Venct. " & var_datavencimento & ", validade " & vValidadeBoleto & ". Valor R$ "& var_valordocumento &" . Impressão automática.', @FUNC_ID = " & Session("FUNC_ID"), vFilialDestino
		else
			BD.Execute "EXEC prc_CriaAndamentoCobranca @CTRA_ID = " & rsParcelas("CTRA_ID") & ", @ANCO_DataHora = '" & vHora & "', @STAC_ID = 23, @TDEV_ID = NULL, @ANCO_Descricao = 'Emissão do boleto " & rsParcelas("PARC_NumDocumento") & ", referente à parcela " & rsParcelas("PARC_Numero") & ". Venct. " & var_datavencimento & ", validade " & vValidadeBoleto & ". Impressão automática.', @FUNC_ID = " & Session("FUNC_ID")
			PreencheLOG "EXEC prc_CriaAndamentoCobranca @CTRA_ID = " & rsParcelas("CTRA_ID") & ", @ANCO_DataHora = '" & vHora & "', @STAC_ID = 23, @TDEV_ID = NULL, @ANCO_Descricao = 'Emissão do boleto " & rsParcelas("PARC_NumDocumento") & ", referente à parcela " & rsParcelas("PARC_Numero") & ". Venct. " & var_datavencimento & ", validade " & vValidadeBoleto & ". Impressão automática.', @FUNC_ID = " & Session("FUNC_ID"), vFilialDestino
		end if
		
		'end if
	elseif Request.QueryString("atraso") <> "" then
		if vCont_id = 43 then
			var_contrato = rsParcelas("CTRA_Conta") & rsParcelas("CTRA_Numero") & " - " & rsEmpresa("CONT_Fantasia")
		elseif rsParcelas("CART_ID") = 85 then
			Dim rsNumTituloCt2, strDescTituloCt2
			Set rsNumTituloCt2 = Bd.Execute("SELECT * FROM Transacoes_de_Acordo t WITH (NOLOCK) LEFT JOIN Tipos_de_Documento td WITH (NOLOCK) ON t.TDOC_ID = td.TDOC_ID WHERE ACOR_ID = " & rsParcelas("ACOR_ID"))
			if rsNumTituloCt2.EOF then
				Set rsNumTituloCt2 = Bd.Execute("SELECT * FROM Transacoes t WITH (NOLOCK) LEFT JOIN Tipos_de_Documento td WITH (NOLOCK) ON t.TDOC_ID = td.TDOC_ID WHERE CTRA_ID = " & rsParcelas("CTRA_ID"))
			end if
			strDescTituloCt2 = ""
			Do while not rsNumTituloCt2.EOF
				strDescTituloCt2 = strDescTituloCt2 & rsNumTituloCt2("TRAN_NumTitulo") & ", "	
				rsNumTituloCt2.movenext
			loop
			strDescTituloCt2 = Left(strDescTituloCt2, len(strDescTituloCt2) - 2)
			var_contrato = strDescTituloCt2			

		else
			var_contrato = FormataContrato(rsParcelas("CTRA_Numero")) & " - " & rsEmpresa("CONT_Fantasia")
		end if
		var_parcelaplano = rsParcelas("PARC_Numero") & " / " & rsParcelas("Plano")

		vValorTaxaBoleto = rsParcelas("BOAT_TaxaBoleto")

		var_datavencimento = rsParcelas("PARC_Vencimento") '"05/06/2002"
		var_valordocumento = FormatNumber(rsParcelas("BOAT_TaxaBoleto") + rsParcelas("PARC_ValorTotal")) '"70,00"

		if Session("CodigoCliente") = 13 and vCONT_ID = 11 then
			vTel = "(21)3212-3771"
		elseif Session("CodigoCliente") = 13 and (rsParcelas("CART_ID") = 38001 or rsParcelas("CART_ID") = 47001 or rsParcelas("CART_ID") = 42001 or rsParcelas("CART_ID") = 48001) then
			vTel = "(21)3212-3792"
		elseif Session("CodigoCliente") = 13 and vCONT_ID = 27 then
			vTel = "(21)3212-3780"
		elseif Session("CodigoCliente") = 13 and vCONT_ID = 14 then
			vTel = "(21)3212-3788"
		elseif Session("CodigoCliente") = 13 and (vCONT_ID = 16 or vCONT_ID = 17 or vCONT_ID = 61) then
			vTel = "(21)3212-3789"
		elseif Session("CodigoCliente") = 13 and vCONT_ID = 77 then
			vTel = "(21)3212-3793"
		elseif Session("CodigoCliente") = 13 and vCONT_ID = 80 then
			vTel = "(21)3212-3794"
		elseif Session("CodigoCliente") = 13 and vCONT_ID = 86 then
			vTel = "(21)3212-3796"
		elseif Session("CodigoCliente") = 13 and vCONT_ID = 48 then
			vTel = "(21)3212-3786"
		elseif Session("CodigoCliente") = 1 and vCONT_ID = 85 then
			vTel = "(71)3483-1250"
		elseif Session("CodigoCliente") = 23 and vCONT_ID = 58 then
			vTel = "(21)3984-2705"
		elseif Session("CodigoCliente") = 25 and vCONT_ID = 54 then
			vTel = "São Paulo (11) 2165-9423 - Demais localidades 0800 724 23 23"
		elseif Session("CodigoCliente") = 25 and (vCONT_ID = 100 or vCONT_ID = 104) then
			vTel = "SP e Gde. SP: (11) 2165-9436 - Outras Localidades:  0800-7241100"			
		end if

		if vTel <> "" then
			vTel = " Tel.: " & vTel
		end if
		if vFax <> "" then
			vFax = " Fax.: " & vFax
		end if

		vValidadeBoleto = rsParcelas("BOAV_Validade")
		if vCont_id = 1 then
			if CDbl(vValorTaxaBoleto) > 0 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & vValidadeBoleto & ". <BR>Após " & vValidadeBoleto & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "). <BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " <br>O não pagamento até a data de vencimento acarretará no cancelamento do acordo e a inclusão do seu nome no órgão proteção ao crédito de sua região e no SERASA. Implicando na perda do desconto porventura concedido.</b>" 
			else
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & vValidadeBoleto & ". <BR>Após " & vValidadeBoleto & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "). <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " <br>O não pagamento até a data de vencimento acarretará no cancelamento do acordo e a inclusão do seu nome no órgão proteção ao crédito de sua região e no SERASA. Implicando na perda do desconto porventura concedido.</b>" 
			end if
		elseif vCont_id = 8 OR vCont_id = 9 then
			if CDbl(vValorTaxaBoleto) > 0 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & vValidadeBoleto & ". <BR>Após " & vValidadeBoleto & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "). <BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " <br>O não pagamento até a data de vencimento acarretará no cancelamento do acordo e a inclusão do seu nome no órgão proteção ao crédito de sua região.</b>" 
			else
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & vValidadeBoleto & ". <BR>Após " & vValidadeBoleto & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "). <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " <br>O não pagamento até a data de vencimento acarretará no cancelamento do acordo e a inclusão do seu nome no órgão proteção ao crédito de sua região.</b>" 
			end if
		elseif vCont_id = 2 then
			var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & vValidadeBoleto & ". <BR>Após o vencimento cobrar 2% de multa. <BR>Após " & vValidadeBoleto & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " </b>" 
		elseif vCont_id = 43 then
			var_instrucoes = "<B>Pagável em qualquer banco até o dia " & vValidadeBoleto & ". <BR>Após " & vValidadeBoleto & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>CGC/CPF: " & FormataCPFCNPJ(rsParcelas("DEVE_CGCCPF")) & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & ".</b>" 
		elseif vCont_id = 11 and rsParcelas("Plano") = 1 then
			if CDbl(vValorTaxaBoleto) > 0 then
				var_instrucoes = "<BR> <B>Pagável apenas até o vencimento.<BR>Após o vencimento, entrar em contato com a " & cons_cedente2 & " - " & vTel & " " & vFax & ".<BR> Para pagamento em cheque, a quitação deste título ocorrerá após a quitação. <BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & "</b>" 
			else
				var_instrucoes = "<BR> <B>Pagável apenas até o vencimento. <BR>Após o vencimento, entrar em contato com a " & cons_cedente2 & " - " & vTel & " " & vFax & ".<BR> Para pagamento em cheque, a quitação deste título ocorrerá após a quitação. <BR>Referente ao contrato nº: " & var_contrato & "</b>" 
			end if
		elseif vCont_id = 11 and rsParcelas("Plano") > 1 then
			if CDbl(vValorTaxaBoleto) > 0 then
				var_instrucoes = "<BR> <B>Após o vencimento, pagável somente no Banco Bradesco até o 10º dia de vencido. <br>Em caso de atraso, atualizar o valor desde a data de vencimento aplicando TR + 2,00% AM pro rata. <br>Após o 10º dia de vencido, entrar em contato com a " & cons_cedente2 & " - " & vTel & " " & vFax & ". <br> Para pagamento em cheque, a quitação deste título ocorrerá após a compensação.<BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " </b>" 
			else
				var_instrucoes = "<BR> <B>Após o vencimento, pagável somente no Banco Bradesco até o 10º dia de vencido. <br>Em caso de atraso, atualizar o valor desde a data de vencimento aplicando TR + 2,00% AM pro rata. <br>Após o 10º dia de vencido, entrar em contato com a " & cons_cedente2 & " - " & vTel & " " & vFax & ". <br> Para pagamento em cheque, a quitação deste título ocorrerá após a compensação.<BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " </b>" 
			end if
		elseif vCont_id = 14 and Session("CodigoCliente") = 17 then
			var_instrucoes = "<BR> <B>Não receber após o vencimento.<br>Não receber valor inferior ao do documento.<br>Após " & vValidadeBoleto & " entrar em contato com " & cons_cedente2 & " - " & vTel & " " & vFax & ". <br>Empresa mandatária: Cobresp.<BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " </b>" 
		elseif vCont_id = 14 then
			if CDbl(vValorTaxaBoleto) > 0 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & vValidadeBoleto & ". <BR>Após " & vValidadeBoleto & " entrar em contato com " & cons_cedente2 & " - " & vTel & " " & vFax & ". <BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " </b>" 
			else
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & vValidadeBoleto & ". <BR>Após " & vValidadeBoleto & " entrar em contato com " & cons_cedente2 & " - " & vTel & " " & vFax & ". <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " </b>" 
			end if
		elseif vCont_id = 48 then
			if CDbl(vValorTaxaBoleto) > 0 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & vValidadeBoleto & ". <BR>Após " & vValidadeBoleto & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & ").<BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " </b>. O não pagamento até a data de vencimento acarretará no cancelamento do acordo e a inclusão do seu nome no órgão de proteção ao crédito de sua região. " 
			else
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & vValidadeBoleto & ". <BR>Após " & vValidadeBoleto & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & ").<BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " </b>. O não pagamento até a data de vencimento acarretará no cancelamento do acordo e a inclusão do seu nome no órgão de proteção ao crédito de sua região. " 
			end if
		elseif vCont_id = 3 then

				vContratos = ""
						Set rsOutrosContratos = BD.Execute("SELECT TRAN_NumTitulo FROM Transacoes WHERE CTRA_ID = " & rsParcelas("CTRA_ID") & " AND TDOC_ID <> 307")
						do While not rsOutrosContratos.EOF 
							vContratos = vContratos & rsOutrosContratos("TRAN_NumTitulo") & ", "
							rsOutrosContratos.MoveNext
						loop
				vContratos = left(vContratos, len(vContratos) - 2)
		
			if CDbl(vValorTaxaBoleto) > 0 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & vValidadeBoleto & ". <BR>Após " & vValidadeBoleto & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao(s) contrato(s) nº: " & vContratos & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " - CNPJ/CPF: " & FormataCPFCNPJ(rsParcelas("DEVE_CGCCPF")) & " <br>O envio desse boleto não inibe uma possível Ação de Débito efetuado pelo HSBC na conta corrente.</b>" 
			else
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & vValidadeBoleto & ". <BR>Após " & vValidadeBoleto & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Referente ao(s) contrato(s) nº: " & vContratos & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " - CNPJ/CPF: " & FormataCPFCNPJ(rsParcelas("DEVE_CGCCPF")) & "<br>O envio desse boleto não inibe uma possível Ação de Débito efetuado pelo HSBC na conta corrente.</b>" 
			end if
		elseif vCont_id = 4 then
			var_contrato = Mid(var_contrato,1 ,2) & "**.****.****." & Mid(var_contrato,16 ,4)
			if CDbl(vValorTaxaBoleto) > 0 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & vValidadeBoleto & ". <BR>Após " & vValidadeBoleto & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao(s) contrato(s) nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " - CNPJ/CPF: " & FormataCPFCNPJ(rsParcelas("DEVE_CGCCPF")) & " <br>O envio desse boleto não inibe uma possível Ação de Débito efetuado pelo HSBC na conta corrente.</b>" 
			else
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & vValidadeBoleto & ". <BR>Após " & vValidadeBoleto & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " - CNPJ/CPF: " & FormataCPFCNPJ(rsParcelas("DEVE_CGCCPF")) & "<br>O envio desse boleto não inibe uma possível Ação de Débito efetuado pelo HSBC na conta corrente.</b>" 
			end if
		elseif vCont_id = 54 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & vValidadeBoleto & ". <BR>Após " & vValidadeBoleto & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " </b> " 
		elseif vCONT_id = 53 then
			if CDbl(vValorTaxaBoleto) > 0 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & ". <BR>Após " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao saldo remanescente do contrato: "& var_contrato &" - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") &". </b>" 
			else
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & ". <BR>Após " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Referente ao saldo remanescente do contrato: "& var_contrato &" - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & ". </b>" 
			end if

		elseif rsPol("PNEG_CorrigeParcelaAtraso") then
			var_instrucoes = "<BR> <B>Senhor Caixa: Cobrar " & FormatCurrency(AtualizaValorParcelaBoleto(rsParcelas("CART_ID"), CDate(rsParcelas("PARC_Vencimento")), CDate(rsParcelas("PARC_Vencimento")) + 1, rsParcelas("PARC_ValorTotal"), rsParcelas("FNEG_ID")) - rsParcelas("PARC_ValorTotal")) & " por dia de atraso.<BR>Pagável em qualquer banco até o dia " & vValidadeBoleto & " acrescido dos encargos por dia de atraso. <BR>Após " & vValidadeBoleto & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " </b>" 
			vJurosAtraso = true
		else
			if CDbl(vValorTaxaBoleto) > 0 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & vValidadeBoleto & ". <BR>Após " & vValidadeBoleto & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " </b>" 
			else
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & vValidadeBoleto & ". <BR>Após " & vValidadeBoleto & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " </b>" 
			end if
		end if
		var_textoparcela = rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano")
		
		'Preenche o andamento da cobrança
		'if Session("CodigoCliente") <> 0 then
			BD.Execute "EXEC prc_CriaAndamentoCobranca @CTRA_ID = " & rsParcelas("CTRA_ID") & ", @ANCO_DataHora = '" & vHora & "', @STAC_ID = 23, @TDEV_ID = NULL, @ANCO_Descricao = 'Emissão do boleto " & rsParcelas("PARC_NumDocumento") & ", referente à parcela " & rsParcelas("PARC_Numero") & ", vencimento " & rsParcelas("PARC_Vencimento") & ", valor " & FormatCurrency(rsParcelas("PARC_ValorTotal")) & ", taxa de boleto " & FormatCurrency(rsParcelas("BOAT_TaxaBoleto")) & ". Impressão de boletos de parcela em atraso.', @FUNC_ID = " & Session("FUNC_ID")
			PreencheLOG "EXEC prc_CriaAndamentoCobranca @CTRA_ID = " & rsParcelas("CTRA_ID") & ", @ANCO_DataHora = '" & vHora & "', @STAC_ID = 23, @TDEV_ID = NULL, @ANCO_Descricao = 'Emissão do boleto " & rsParcelas("PARC_NumDocumento") & ", referente à parcela " & rsParcelas("PARC_Numero") & ", vencimento " & rsParcelas("PARC_Vencimento") & ", valor " & FormatCurrency(rsParcelas("PARC_ValorTotal")) & ", taxa de boleto " & FormatCurrency(rsParcelas("BOAT_TaxaBoleto")) & ". Impressão de boletos de parcela em atraso.', @FUNC_ID = " & Session("FUNC_ID"), vFilialDestino
		'end if
	elseif Request.QueryString("avulso") <> "" then
		if vCont_id = 43 then
			var_contrato = rsParcelas("CTRA_Conta") & rsParcelas("CTRA_Numero") & " - " & rsBanco("CONT_Fantasia")
		elseif rsParcelas("CART_ID") = 85 then
			Dim rsNumTituloCt1, strDescTituloCt1
			Set rsNumTituloCt1 = Bd.Execute("SELECT * FROM Transacoes t WITH (NOLOCK) LEFT JOIN Tipos_de_Documento td WITH (NOLOCK) ON t.TDOC_ID = td.TDOC_ID WHERE CTRA_ID = " & rsParcelas("CTRA_ID"))
			strDescTituloCt1 = ""
			Do while not rsNumTituloCt1.EOF
				strDescTituloCt1 = strDescTituloCt1 & rsNumTituloCt1("TRAN_NumTitulo") & ", "	
				rsNumTituloCt1.movenext
			loop
			strDescTituloCt1 = Left(strDescTituloCt1, len(strDescTituloCt1) - 2)
			var_contrato = strDescTituloCt1			
		else
			var_contrato = FormataContrato(rsParcelas("CTRA_Numero")) & " - " & rsBanco("CONT_Fantasia")
		end if
		var_parcelaplano = "1 / " & rsParcelas("Plano")
		
		vValorTaxaBoleto = rsParcelas("BOAV_TaxaBoleto")

		var_datavencimento = rsParcelas("BOAV_Vencimento") '"05/06/2002"
		var_valordocumento = FormatNumber(CDbl(vValorTaxaBoleto) + CDbl(rsParcelas("BOAV_Valor"))) '"70,00"

		if Session("CodigoCliente") = 13 and vCONT_ID = 11 then
			vTel = "(21)3212-3771"
		elseif Session("CodigoCliente") = 13 and (rsParcelas("CART_ID") = 38001 or rsParcelas("CART_ID") = 47001 or rsParcelas("CART_ID") = 42001 or rsParcelas("CART_ID") = 48001) then
			vTel = "(21)3212-3792"
		elseif Session("CodigoCliente") = 13 and vCONT_ID = 27 then
			vTel = "(21)3212-3780"
		elseif Session("CodigoCliente") = 13 and vCONT_ID = 14 then
			vTel = "(21)3212-3788"
		elseif Session("CodigoCliente") = 13 and (vCONT_ID = 16 or vCONT_ID = 17 or vCONT_ID = 61) then
			vTel = "(21)3212-3789"
		elseif Session("CodigoCliente") = 13 and vCONT_ID = 77 then
			vTel = "(21)3212-3793"
		elseif Session("CodigoCliente") = 13 and vCONT_ID = 80 then
			vTel = "(21)3212-3794"
		elseif Session("CodigoCliente") = 13 and vCONT_ID = 86 then
			vTel = "(21)3212-3796"
		elseif Session("CodigoCliente") = 13 and vCONT_ID = 48 then
			vTel = "(21)3212-3786"
		elseif Session("CodigoCliente") = 1 and vCONT_ID = 85 then
			vTel = "(71)3483-1250"
		elseif Session("CodigoCliente") = 23 and vCONT_ID = 58 then
			vTel = "(21)3984-2705"
		elseif Session("CodigoCliente") = 25 and vCONT_ID = 54 then
			vTel = "São Paulo (11) 2165-9423 - Demais localidades 0800 724 23 23"
		elseif Session("CodigoCliente") = 25 and (vCONT_ID = 100 or vCONT_ID = 104) then
			vTel = "SP e Gde. SP: (11) 2165-9436 - Outras Localidades:  0800-7241100"			
		end if

		if vTel <> "" then
			vTel = " Tel.: " & vTel
		end if
		if vFax <> "" then
			vFax = " Fax.: " & vFax
		end if

		vValidadeBoleto = rsParcelas("BOAV_Validade")
		if vCont_id = 1 then
			if CDbl(rsParcelas("BOAV_TaxaBoleto")) > 0 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & rsParcelas("BOAV_Validade") & ". <BR>Após " & rsParcelas("BOAV_Validade") & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & ")<BR>Valor do documento acrescido de " & FormatCurrency(rsParcelas("BOAV_TaxaBoleto")) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " <br>O não pagamento até a data de vencimento acarretará no cancelamento do acordo e a inclusão do seu nome no órgão proteção ao crédito de sua região e no SERASA. Implicando na perda do desconto porventura concedido.</b>" 
			else
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & rsParcelas("BOAV_Validade") & ". <BR>Após " & rsParcelas("BOAV_Validade") & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & ")<BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " <br>O não pagamento até a data de vencimento acarretará no cancelamento do acordo e a inclusão do seu nome no órgão proteção ao crédito de sua região e no SERASA. Implicando na perda do desconto porventura concedido.</b>" 
			end if
		elseif vCont_id = 8 OR vCont_id = 9 then
			if CDbl(rsParcelas("BOAV_TaxaBoleto")) > 0 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & rsParcelas("BOAV_Validade") & ". <BR>Após " & rsParcelas("BOAV_Validade") & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & ")<BR>Valor do documento acrescido de " & FormatCurrency(rsParcelas("BOAV_TaxaBoleto")) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " <br>O não pagamento até a data de vencimento acarretará no cancelamento do acordo e a inclusão do seu nome no órgão proteção ao crédito de sua região.</b>" 
			else
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & rsParcelas("BOAV_Validade") & ". <BR>Após " & rsParcelas("BOAV_Validade") & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & ")<BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " <br>O não pagamento até a data de vencimento acarretará no cancelamento do acordo e a inclusão do seu nome no órgão proteção ao crédito de sua região.</b>" 
			end if
		elseif vCont_id = 2 then
			var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & rsParcelas("BOAV_Validade") & ". <BR>Após o vencimento cobrar 2% de multa. <BR>Após " & rsParcelas("BOAV_Validade") & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " </b>" 
		elseif vCont_id = 43 then
		
			Dim rsNumTitulo2, strDescTitulo2
			Set rsNumTitulo2 = Bd.Execute("SELECT * FROM Transacoes t WITH (NOLOCK) LEFT JOIN Tipos_de_Documento td WITH (NOLOCK) ON t.TDOC_ID = td.TDOC_ID WHERE CTRA_ID = " & rsParcelas("CTRA_ID"))
			strDescTitulo2 = ""
			Do while not rsNumTitulo2.EOF
				strDescTitulo2 = strDescTitulo2 & rsNumTitulo2("TRAN_NumTitulo") & ", "	
				rsNumTitulo2.movenext
			loop
		
			strDescTitulo2 = Left(strDescTitulo2, len(strDescTitulo2) - 2)
		
			var_instrucoes = "<B>Pagável em qualquer banco até o dia " & rsParcelas("BOAV_Validade") & ". <BR>Após " & rsParcelas("BOAV_Validade") & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>CGC/CPF: " & FormataCPFCNPJ(rsParcelas("DEVE_CGCCPF")) & " - Parcela/Plano: " & var_parcelaplano & ".<br></b>Titulo(s):&nbsp;" & strDescTitulo2 & "&nbsp;<input type=text name=Obs size=70 class=texto2><br>Sr. Cliente Caso tenha ocorrido mudança de endereço, favor comparecer agência para atualização dos dados cadastrais para que o carnê com os próximos pagamentos chegue no endereço correto." 

		elseif vCont_id = 3 then
			vQtdContratosHSBC = 0
			'if rsParcelas("Plano") > 1 then
				'vContratos = FormataContrato(rsParcelas("CTRA_Numero")) & ", "
				vContratos = ""
				'if DateDiff("d", rsParcelas("CTRA_VencDebito"), date) > 60 and DateDiff("d", rsParcelas("CTRA_VencDebito"), date) < 180 then
				'if DateDiff("d", rsParcelas("CTRA_VencDebito"), date) > 60 then
					'Set rsOutrosContratos = BD.Execute("SELECT SUM(TRAN_Valor) Soma FROM Transacoes t WITH (NOLOCK) JOIN Contratos c WITH (NOLOCK) ON t.CTRA_ID = c.CTRA_ID JOIN Devedores d WITH (NOLOCK) ON c.DEVE_ID = d.DEVE_ID JOIN Carteiras ca WITH (NOLOCK) ON c.CART_ID = ca.CART_ID WHERE DEVE_CGCCPF = '" & rsParcelas("DEVE_CGCCPF") & "' AND DateDiff(day, CTRA_VencDebito, getdate()) < 240 AND CONT_ID = 3 AND SCON_ID NOT IN (2, 4, 7)")
					'Set rsOutrosContratos = BD.Execute("SELECT SUM(TRAN_Valor) Soma FROM Transacoes t WITH (NOLOCK) JOIN Contratos c WITH (NOLOCK) ON t.CTRA_ID = c.CTRA_ID JOIN Devedores d WITH (NOLOCK) ON c.DEVE_ID = d.DEVE_ID JOIN Carteiras ca WITH (NOLOCK) ON c.CART_ID = ca.CART_ID WHERE DEVE_CGCCPF = '" & rsParcelas("DEVE_CGCCPF") & "' AND CONT_ID = 3 AND SCON_ID NOT IN (2, 4, 7)")
					'if rsOutrosContratos("Soma") < 15000 then
						'Set rsOutrosContratos = BD.Execute("SELECT CTRA_Numero FROM Contratos c WITH (NOLOCK) JOIN Devedores d WITH (NOLOCK) ON c.DEVE_ID = d.DEVE_ID JOIN Carteiras ca WITH (NOLOCK) ON c.CART_ID = ca.CART_ID WHERE DEVE_CGCCPF = '" & rsParcelas("DEVE_CGCCPF") & "' AND CTRA_ID <> " & rsParcelas("ctra_id") & " AND DateDiff(day, CTRA_VencDebito, getdate()) < 240 AND CONT_ID = 3 AND SCON_ID NOT IN (2, 4, 7)")
						Set rsOutrosContratos = BD.Execute("SELECT TRAN_NumTitulo FROM Transacoes WHERE CTRA_ID = " & rsParcelas("CTRA_ID") & " AND TDOC_ID <> 307")
						do While not rsOutrosContratos.EOF 
							vContratos = vContratos & rsOutrosContratos("TRAN_NumTitulo") & ", "
							rsOutrosContratos.MoveNext
							vQtdContratosHSBC = vQtdContratosHSBC + 1
						loop
					'end if
				'end if
				vContratos = left(vContratos, len(vContratos) - 2)
				vContratos = vContratos & " - " & rsBanco("CONT_Fantasia")
				if vQtdContratosHSBC > 0 then
					var_contrato = "&nbsp;"
				end if
			'else
			'	vContratos = var_contrato
			'end if
			var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & rsParcelas("BOAV_Validade") & ". <BR>Após " & rsParcelas("BOAV_Validade") & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Referente ao(s) contrato(s) nº: " & vContratos & "<br>Parcela/Plano: " & var_parcelaplano & " - CNPJ/CPF: " & FormataCPFCNPJ(rsParcelas("DEVE_CGCCPF")) & " <br>O envio desse boleto não inibe uma possível Ação de Débito efetuado pelo HSBC na conta corrente.</b>" 
		elseif vCont_id = 4 then
			'HSBC CARTÂO
			var_contrato = Mid(var_contrato,1 ,2) & "**.****.****." & Mid(var_contrato,16 ,4)						
			var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & rsParcelas("BOAV_Validade") & ". <BR>Após " & rsParcelas("BOAV_Validade") & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Referente ao contrato nº: " & var_contrato & "<br>Parcela/Plano: " & var_parcelaplano & " - CNPJ/CPF: " & FormataCPFCNPJ(rsParcelas("DEVE_CGCCPF")) & " <br>O envio desse boleto não inibe uma possível Ação de Débito efetuado pelo HSBC na conta corrente.</b>" 
		elseif vCont_id = 11 and rsParcelas("Plano") = 1 then
			if CDbl(rsParcelas("BOAV_TaxaBoleto")) > 0 then
				var_instrucoes = "<BR> <B>Pagável apenas até o vencimento.<BR>Após o vencimento, entrar em contato com a " & cons_cedente2 & " - " & vTel & " " & vFax & ".<BR> Para pagamento em cheque, a quitação deste título ocorrerá após a quitação. <BR>Valor do documento acrescido de " & FormatCurrency(rsParcelas("BOAV_TaxaBoleto")) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & "</b>" 
			else
				var_instrucoes = "<BR> <B>Pagável apenas até o vencimento. <BR>Após o vencimento, entrar em contato com a " & cons_cedente2 & " - " & vTel & " " & vFax & ".<BR> Para pagamento em cheque, a quitação deste título ocorrerá após a quitação. <BR>Referente ao contrato nº: " & var_contrato & "</b>" 
			end if
		elseif vCont_id = 11 and rsParcelas("Plano") > 1 then
			if CDbl(rsParcelas("BOAV_TaxaBoleto")) > 0 then
				var_instrucoes = "<BR> <B>Após o vencimento, pagável somente no Banco Bradesco até o 10º dia de vencido. <br>Em caso de atraso, atualizar o valor desde a data de vencimento aplicando TR + 2,00% AM pro rata. <br>Após o 10º dia de vencido, entrar em contato com a " & cons_cedente2 & " - " & vTel & " " & vFax & ". <br> Para pagamento em cheque, a quitação deste título ocorrerá após a compensação.<BR>Valor do documento acrescido de " & FormatCurrency(rsParcelas("BOAV_TaxaBoleto")) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " </b> - CNPJ/CPF: " & FormataCPFCNPJ(rsParcelas("DEVE_CGCCPF")) & "" 
			else
				var_instrucoes = "<BR> <B>Após o vencimento, pagável somente no Banco Bradesco até o 10º dia de vencido. <br>Em caso de atraso, atualizar o valor desde a data de vencimento aplicando TR + 2,00% AM pro rata. <br>Após o 10º dia de vencido, entrar em contato com a " & cons_cedente2 & " - " & vTel & " " & vFax & ". <br> Para pagamento em cheque, a quitação deste título ocorrerá após a compensação.<BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " </b>" 
			end if
		elseif vCont_id = 14 and Session("CodigoCliente") = 17 then
			var_instrucoes = "<BR> <B>Não receber após o vencimento.<br>Não receber valor inferior ao do documento.<br>Após " & rsParcelas("BOAV_Validade") & " entrar em contato com " & cons_cedente2 & " - " & vTel & " " & vFax & ". <br>Empresa mandatária: Cobresp.<BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " </b>" 
		elseif vCont_id = 14 then
			if CDbl(rsParcelas("BOAV_TaxaBoleto")) > 0 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & rsParcelas("BOAV_Validade") & ". <BR>Após " & rsParcelas("BOAV_Validade") & " entrar em contato com " & cons_cedente2 & " - " & vTel & " " & vFax & ". <BR>Valor do documento acrescido de " & FormatCurrency(rsParcelas("BOAV_TaxaBoleto")) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " </b>" 
			else
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & rsParcelas("BOAV_Validade") & ". <BR>Após " & rsParcelas("BOAV_Validade") & " entrar em contato com " & cons_cedente2 & " - " & vTel & " " & vFax & ". <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " </b>" 
			end if
		elseif vCont_id = 48 then
			if CDbl(vValorTaxaBoleto) > 0 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & rsParcelas("BOAV_Validade") & ". <BR>Após " & rsParcelas("BOAV_Validade") & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & ").<BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " </b>. O não pagamento até a data de vencimento acarretará no cancelamento do acordo e a inclusão do seu nome no órgão de proteção ao crédito de sua região. " 
			else
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & rsParcelas("BOAV_Validade") & ". <BR>Após " & rsParcelas("BOAV_Validade") & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & ").<BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " </b>. O não pagamento até a data de vencimento acarretará no cancelamento do acordo e a inclusão do seu nome no órgão de proteção ao crédito de sua região. " 
			end if
		elseif vCont_id = 58 then
			if CDbl(vValorTaxaBoleto) > 0 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & vValidadeBoleto & ". <BR>Após " & vValidadeBoleto & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR> Referente ao pagto de parcela do contrato nº: " & var_contrato & " </b>" 
			else
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & vValidadeBoleto & ". <BR>Após " & vValidadeBoleto & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Referente ao pagto de parcela do contrato nº: " & var_contrato & " </b>" 
			end if
		elseif vCont_id = 71 then
				
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & rsParcelas("BOAV_Validade") & ". <BR>Após " & rsParcelas("BOAV_Validade") & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " </b>" 
								
				if date() >= cdate("23/11/2005") and date() <= cdate("23/12/2005") and var_parcelaplano = "1 / 1" then
					var_instrucoes = var_instrucoes & "<br> Seu Guanabara Card será liberado para compras após quitação e processamento total do pagamento.<br> Para isso, após quitação, entre em contato pelo telefone " & vTel & " <br> ATENÇÃO: ESTA PROMOÇÃO É VÁLIDA SOMENTE ATÉ 23/12/2005. APROVEITE!"		
				end if

		elseif vCont_id = 80 then

			Set rsChequeInfo = Bd.Execute("SELECT DADO_Valor FROM Dados_Adicionais_do_Contrato WHERE CTRA_ID = " & rsParcelas("CTRA_ID") & " AND TDAD_ID = 87")
			
			var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & rsParcelas("BOAV_Validade") & ". <BR>Após " & rsParcelas("BOAV_Validade") & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " </b>" 
							
			if not rsChequeInfo.EOF then
				var_instrucoes = var_instrucoes & " - " & rsChequeInfo("DADO_Valor") 
			end if 
		
		elseif vCont_id = 54 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & rsParcelas("BOAV_Validade") & ". <BR>Após " & rsParcelas("BOAV_Validade") & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " </b>" 

		elseif vCONT_id = 53 then
			if CDbl(vValorTaxaBoleto) > 0 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & ". <BR>Após " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao saldo remanescente do contrato: "& var_contrato &" - Parcela/Plano: " & var_parcelaplano &". </b>" 
			else
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & ". <BR>Após " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Referente ao saldo remanescente do contrato: "& var_contrato &" - Parcela/Plano: " & var_parcelaplano & ". </b>" 
			end if
		else
			if CDbl(vValorTaxaBoleto) > 0 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & rsParcelas("BOAV_Validade") & ". <BR>Após " & rsParcelas("BOAV_Validade") & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " </b>" 
			else
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & rsParcelas("BOAV_Validade") & ". <BR>Após " & rsParcelas("BOAV_Validade") & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " </b>" 
			end if
		end if
		var_textoparcela = "1 / " & rsParcelas("Plano")

		'Preenche o andamento da cobrança
		'if Session("CodigoCliente") <> 0 then
		BD.Execute "EXEC prc_CriaAndamentoCobranca @CTRA_ID = " & rsParcelas("ctra_id") & ", @ANCO_DataHora = '" & vHora & "', @STAC_ID = 23, @TDEV_ID = NULL, @ANCO_Descricao = 'Emissão do boleto avulso " & var_nossonumero & ", no valor de R$ " & var_valordocumento & ", com vencimento em " & var_datavencimento & ", referente à parcela/plano " & var_parcelaplano & "', @FUNC_ID = " & Session("FUNC_ID")
		PreencheLOG "EXEC prc_CriaAndamentoCobranca @CTRA_ID = " & rsParcelas("ctra_id") & ", @ANCO_DataHora = '" & vHora & "', @STAC_ID = 23, @TDEV_ID = NULL, @ANCO_Descricao = 'Emissão do boleto avulso " & var_nossonumero & ", no valor de R$ " & var_valordocumento & ", com vencimento em " & var_datavencimento & ", referente à parcela/plano " & var_parcelaplano & "', @FUNC_ID = " & Session("FUNC_ID"), vFilialDestino
		'end if		
		if rsParcelas("SCON_ID") <> 2 then
			vHora = DateAdd("s", 1, vHora)
			strSQL = "EXEC prc_CriaAndamentoCobranca @CTRA_ID = " & rsParcelas("ctra_id") & ", @ANCO_DataHora = '" & vHora & "', @STAC_ID = 13, @TDEV_ID = NULL, @ANCO_Descricao = 'Contrato direcionado para a fila de cobrança Pré-Acordo. Motivo: Boleto avulso emitido.', @FUNC_ID = 0"
			BD.Execute strSQL
			PreencheLOG strSQL, vFilialDestino

			BD.Execute "prc_INS_HistoricoAgendamento_Forcado 3001, '" & now & "', " & rsParcelas("ctra_id") & ", " & Session("FILI_ID") & ", " & Session("FUNC_ID")
			PreencheLOG "prc_INS_HistoricoAgendamento_Forcado 3001, '" & now & "', " & rsParcelas("ctra_id") & ", " & Session("FILI_ID") & ", " & Session("FUNC_ID"), vFilialDestino
		end if
	elseif vTipoBoleto = 1 then
		if vCont_id = 43 then
			var_contrato = rsParcelas("CTRA_Conta") & rsParcelas("CTRA_Numero") & " - " & rsBanco("CONT_Fantasia")
		elseif rsParcelas("CART_ID") = 85 then
			Dim rsNumTituloCt, strDescTituloCt
			Set rsNumTituloCt = Bd.Execute("SELECT * FROM Transacoes_de_Acordo t WITH (NOLOCK) LEFT JOIN Tipos_de_Documento td WITH (NOLOCK) ON t.TDOC_ID = td.TDOC_ID WHERE ACOR_ID = " & rsParcelas("ACOR_ID"))
			if rsNumTituloCt.EOF then
				Set rsNumTituloCt = Bd.Execute("SELECT * FROM Transacoes t WITH (NOLOCK) LEFT JOIN Tipos_de_Documento td WITH (NOLOCK) ON t.TDOC_ID = td.TDOC_ID WHERE CTRA_ID = " & rsParcelas("CTRA_ID"))
			end if
			strDescTituloCt = ""
			Do while not rsNumTituloCt.EOF
				strDescTituloCt = strDescTituloCt & rsNumTituloCt("TRAN_NumTitulo") & ", "	
				rsNumTituloCt.movenext
			loop
			strDescTituloCt = Left(strDescTituloCt, len(strDescTituloCt) - 2)
			var_contrato = strDescTituloCt
		else
			var_contrato = FormataContrato(rsParcelas("CTRA_Numero")) & " - " & rsBanco("CONT_Fantasia")
		end if
		var_parcelaplano = rsParcelas("PARC_Numero") & " / " & rsParcelas("Plano")

		var_datavencimento = rsParcelas("PARC_Vencimento") '"05/06/2002"
		var_valordocumento = FormatNumber(vValorTaxaBoleto + rsParcelas("PARC_ValorTotal")) '"70,00"

		if Session("CodigoCliente") = 13 and vCONT_ID = 11 then
			vTel = "(21)3212-3771"
		elseif Session("CodigoCliente") = 13 and (rsParcelas("CART_ID") = 38001 or rsParcelas("CART_ID") = 47001 or rsParcelas("CART_ID") = 42001 or rsParcelas("CART_ID") = 48001) then
			vTel = "(21)3212-3792"
		elseif Session("CodigoCliente") = 13 and vCONT_ID = 27 then
			vTel = "(21)3212-3780"
		elseif Session("CodigoCliente") = 13 and vCONT_ID = 14 then
			vTel = "(21)3212-3788"
		elseif Session("CodigoCliente") = 13 and (vCONT_ID = 16 or vCONT_ID = 17 or vCONT_ID = 61) then
			vTel = "(21)3212-3789"
		elseif Session("CodigoCliente") = 13 and vCONT_ID = 77 then
			vTel = "(21)3212-3793"
		elseif Session("CodigoCliente") = 13 and vCONT_ID = 80 then
			vTel = "(21)3212-3794"
		elseif Session("CodigoCliente") = 13 and vCONT_ID = 86 then
			vTel = "(21)3212-3796"
		elseif Session("CodigoCliente") = 13 and vCONT_ID = 48 then
			vTel = "(21)3212-3786"
		elseif Session("CodigoCliente") = 1 and vCONT_ID = 85 then
			vTel = "(71)3483-1250"
		elseif Session("CodigoCliente") = 23 and vCONT_ID = 58 then
			vTel = "(21)3984-2705"
		elseif Session("CodigoCliente") = 25 and vCONT_ID = 54 then
			vTel = "São Paulo (11) 2165-9423 - Demais localidades 0800 724 23 23"
		elseif Session("CodigoCliente") = 25 and (vCONT_ID = 100 or vCONT_ID = 104) then
			vTel = "SP e Gde. SP: (11) 2165-9436 - Outras Localidades:  0800-7241100"			
		end if

		if vTel <> "" then
			vTel = " Tel.: " & vTel
		end if
		if vFax <> "" then
			vFax = " Fax.: " & vFax
		end if

		if Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")) = "" then
			vValidadeBoleto = var_datavencimento
		else
			vValidadeBoleto = DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento)
		end if
		if vCont_id = 1 then
			if CDbl(vValorTaxaBoleto) > 0 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & ". <BR>Após " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "). <BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " <br>O não pagamento até a data de vencimento acarretará no cancelamento do acordo e a inclusão do seu nome no órgão proteção ao crédito de sua região e no SERASA. Implicando na perda do desconto porventura concedido.</b>" 
			else
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & ". <BR>Após " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "). <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " <br>O não pagamento até a data de vencimento acarretará no cancelamento do acordo e a inclusão do seu nome no órgão proteção ao crédito de sua região e no SERASA. Implicando na perda do desconto porventura concedido.</b>" 
			end if
		elseif vCont_id = 8 OR vCont_id = 9 then
			if CDbl(vValorTaxaBoleto) > 0 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & ". <BR>Após " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "). <BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " <br>O não pagamento até a data de vencimento acarretará no cancelamento do acordo e a inclusão do seu nome no órgão proteção ao crédito de sua região.</b>" 
			else
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & ". <BR>Após " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "). <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " <br>O não pagamento até a data de vencimento acarretará no cancelamento do acordo e a inclusão do seu nome no órgão proteção ao crédito de sua região.</b>" 
			end if
		elseif vCont_id = 2 then
			var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & ". <BR>Após o vencimento cobrar 2% de multa. <BR>Após " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " </b>" 
		elseif vCont_id = 43 then

			Dim rsNumTitulo1, strDescTitulo1
			Set rsNumTitulo1 = Bd.Execute("SELECT * FROM Transacoes_de_Acordo t WITH (NOLOCK) LEFT JOIN Tipos_de_Documento td WITH (NOLOCK) ON t.TDOC_ID = td.TDOC_ID WHERE ACOR_ID = " & rsParcelas("ACOR_ID"))
			if rsNumTitulo1.EOF then
				Set rsNumTitulo1 = Bd.Execute("SELECT * FROM Transacoes t WITH (NOLOCK) LEFT JOIN Tipos_de_Documento td WITH (NOLOCK) ON t.TDOC_ID = td.TDOC_ID WHERE CTRA_ID = " & rsParcelas("CTRA_ID"))
			end if
			strDescTitulo1 = ""
			Do while not rsNumTitulo1.EOF
				strDescTitulo1 = strDescTitulo1 & rsNumTitulo1("TRAN_NumTitulo") & ", "	
				rsNumTitulo1.movenext
			loop
		
			strDescTitulo1 = Left(strDescTitulo1, len(strDescTitulo1) - 2)

			var_instrucoes = "<B>Pagável em qualquer banco até o dia " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & ". <BR>Após " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>CNPJ/CPF: " & FormataCPFCNPJ(rsParcelas("DEVE_CGCCPF")) & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & ".<br></b>Titulo(s):&nbsp;" & strDescTitulo1     

		elseif vCont_id = 11 and rsParcelas("Plano") = 1 then
			if vMensagemJuros then
				if CDbl(vValorTaxaBoleto) > 0 then
					var_instrucoes = "<BR> <B>Pagável apenas até o vencimento. <BR>Após o vencimento, entrar em contato com a " & cons_cedente2 & " - " & vTel & " " & vFax & ".<BR> Para pagamento em cheque, a quitação deste título ocorrerá após a quitação. <BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & "</b>" 
				else
					var_instrucoes = "<BR> <B>Pagável apenas até o vencimento. <BR>Após o vencimento, entrar em contato com a " & cons_cedente2 & " - " & vTel & " " & vFax & ".<BR> Para pagamento em cheque, a quitação deste título ocorrerá após a quitação. <BR>Referente ao contrato nº: " & var_contrato & "</b>" 
				end if
			else
				if CDbl(vValorTaxaBoleto) > 0 then
					var_instrucoes = "<BR> <b>Para pagamento em cheque, a quitação deste título ocorrerá após a quitação. <BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & "</b>" 
				else
					var_instrucoes = "<BR> <B>Para pagamento em cheque, a quitação deste título ocorrerá após a quitação. <BR>Referente ao contrato nº: " & var_contrato & "</b>" 
				end if
			end if
		elseif vCont_id = 11 and rsParcelas("Plano") > 1 then
			if vMensagemJuros then
				if CDbl(vValorTaxaBoleto) > 0 then
					var_instrucoes = "<BR> <B>Após o vencimento, pagável somente no Banco Bradesco até o 10º dia de vencido. <br>Em caso de atraso, atualizar o valor desde a data de vencimento aplicando TR + 2,00% AM pro rata. <br>Após o 10º dia de vencido, entrar em contato com a " & cons_cedente2 & " - " & vTel & " " & vFax & ". <br> Para pagamento em cheque, a quitação deste título ocorrerá após a compensação.<BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária.<BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " </b>" 
				else
					var_instrucoes = "<BR> <B>Após o vencimento, pagável somente no Banco Bradesco até o 10º dia de vencido. <br>Em caso de atraso, atualizar o valor desde a data de vencimento aplicando TR + 2,00% AM pro rata. <br>Após o 10º dia de vencido, entrar em contato com a " & cons_cedente2 & " - " & vTel & " " & vFax & ". <br> Para pagamento em cheque, a quitação deste título ocorrerá após a compensação.<BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " </b>" 
				end if
			else
				if CDbl(vValorTaxaBoleto) > 0 then
					var_instrucoes = "<BR> <B>Para pagamento em cheque, a quitação deste título ocorrerá após a compensação.<BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária.<BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " </b>" 
				else
					var_instrucoes = "<BR> <B>Para pagamento em cheque, a quitação deste título ocorrerá após a compensação.<BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " </b>" 
				end if
			end if
		elseif vCont_id = 14 and Session("CodigoCliente") = 17 then
			var_instrucoes = "<BR> <B>Não receber após o vencimento.<br>Não receber valor inferior ao do documento.<br>Após " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & " entrar em contato com " & cons_cedente2 & " - " & vTel & " " & vFax & ". <br>Empresa mandatária: Cobresp.<BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " </b>" 
		elseif vCont_id = 14 then
			if CDbl(rsParcelas("BOAV_TaxaBoleto")) > 0 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & ". <BR>Após " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & " entrar em contato com " & cons_cedente2 & " - " & vTel & " " & vFax & ". <BR>Valor do documento acrescido de " & FormatCurrency(rsParcelas("BOAV_TaxaBoleto")) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " </b>" 
			else
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & ". <BR>Após " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & " entrar em contato com " & cons_cedente2 & " - " & vTel & " " & vFax & ". <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " </b>" 
			end if
		elseif vCont_id = 48 then
			if CDbl(vValorTaxaBoleto) > 0 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & ". <BR>Após " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & ").<BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " </b>. O não pagamento até a data de vencimento acarretará no cancelamento do acordo e a inclusão do seu nome no órgão de proteção ao crédito de sua região. " 
			else
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & ". <BR>Após " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & ").<BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " </b>. O não pagamento até a data de vencimento acarretará no cancelamento do acordo e a inclusão do seu nome no órgão de proteção ao crédito de sua região. " 
			end if
		elseif vCont_id = 3 then

				vContratos = ""
						Set rsOutrosContratos = BD.Execute("SELECT TRAN_NumTitulo FROM Transacoes WHERE CTRA_ID = " & rsParcelas("CTRA_ID") & " AND TDOC_ID <> 307")
						do While not rsOutrosContratos.EOF 
							vContratos = vContratos & rsOutrosContratos("TRAN_NumTitulo") & ", "
							rsOutrosContratos.MoveNext
						loop
				vContratos = left(vContratos, len(vContratos) - 2)
			if CDbl(vValorTaxaBoleto) > 0 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & ". <BR>Após " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao(s) contrato(s) nº: " & vContratos & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " - CNPJ/CPF: " & FormataCPFCNPJ(rsParcelas("DEVE_CGCCPF")) & " <br>O envio desse boleto não inibe uma possível Ação de Débito efetuado pelo HSBC na conta corrente.</b>" 
			else
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & ". <BR>Após " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Referente ao(s) contrato(s) nº: " & vContratos & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " - CNPJ/CPF: " & FormataCPFCNPJ(rsParcelas("DEVE_CGCCPF")) & " <br>O envio desse boleto não inibe uma possível Ação de Débito efetuado pelo HSBC na conta corrente.</b>" 
			end if
		elseif vCont_id = 4 then
		
			if CDbl(vValorTaxaBoleto) > 0 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & ". <BR>Após " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao(s) contrato(s) nº: " & vContratos & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " - CNPJ/CPF: " & FormataCPFCNPJ(rsParcelas("DEVE_CGCCPF")) & " <br>O envio desse boleto não inibe uma possível Ação de Débito efetuado pelo HSBC na conta corrente.</b>" 
			else
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & ". <BR>Após " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Referente ao contrato nº: " & Mid(var_contrato,1 ,2) & "**.****.****." & Mid(var_contrato,16 ,4) & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " - CNPJ/CPF: " & FormataCPFCNPJ(rsParcelas("DEVE_CGCCPF")) & " <br>O envio desse boleto não inibe uma possível Ação de Débito efetuado pelo HSBC na conta corrente.</b>" 
			end if
			var_contrato = Mid(var_contrato,1 ,2) & "**.****.****." & Mid(var_contrato,16 ,4)						
			
		elseif vCONT_id = 58 then
			
			Dim rsNumTitulo, strDescTitulo
			Set rsNumTitulo = Bd.Execute("SELECT * FROM Transacoes_de_Acordo t WITH (NOLOCK) LEFT JOIN Tipos_de_Documento td WITH (NOLOCK) ON t.TDOC_ID = td.TDOC_ID WHERE ACOR_ID = " & rsParcelas("ACOR_ID"))
			if rsNumTitulo.EOF then
				Set rsNumTitulo = Bd.Execute("SELECT * FROM Transacoes t WITH (NOLOCK) LEFT JOIN Tipos_de_Documento td WITH (NOLOCK) ON t.TDOC_ID = td.TDOC_ID WHERE CTRA_ID = " & rsParcelas("CTRA_ID"))
			end if
			strDescTitulo = ""
			Do while not rsNumTitulo.EOF
				strDescTitulo = strDescTitulo & rsNumTitulo("TRAN_NumTitulo") & ", "	
				rsNumTitulo.movenext
			loop
		
			strDescTitulo = Left(strDescTitulo, len(strDescTitulo) - 2)
			
			if CDbl(vValorTaxaBoleto) > 0 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & ". <BR>Após " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " </b>" 
			else
				'var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & ". <BR>Após " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " </b>" 
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & ". <BR>Após " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: "  & strDescTitulo & " </b>" 
			end if
			
		elseif rsPol("PNEG_CorrigeParcelaAtraso") then
			
			if vCont_id = 85 then
	
				var_instrucoes = "<BR>Pagável em qualquer banco até o dia " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & ". <BR>Após " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " </b>" 
			
			else 
			
				var_instrucoes = "<BR> <B>Senhor Caixa: Cobrar " & FormatCurrency(AtualizaValorParcelaBoleto(rsParcelas("CART_ID"), CDate(rsParcelas("PARC_Vencimento")), CDate(rsParcelas("PARC_Vencimento")) + 1, rsParcelas("PARC_ValorTotal"), rsParcelas("FNEG_ID")) - rsParcelas("PARC_ValorTotal")) & " por dia de atraso.<BR>Pagável em qualquer banco até o dia " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & " acrescido dos encargos por dia de atraso. <BR>Após " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " </b>" 
			
			end if
			
			vJurosAtraso = true
		elseif vCONT_id = 54 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & ". <BR>Após " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " </b>" 
		elseif vCONT_id = 53 then
			if CDbl(vValorTaxaBoleto) > 0 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & ". <BR>Após " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao saldo remanescente do contrato: "& var_contrato &" - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") &". </b>" 
			else
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & ". <BR>Após " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Referente ao saldo remanescente do contrato: "& var_contrato &" - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & ". </b>" 
			end if
		else
		
			if CDbl(vValorTaxaBoleto) > 0 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & ". <BR>Após " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " </b>" 
			else
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & ". <BR>Após " & DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano") & " </b>" 
			end if
		end if
		var_textoparcela = rsParcelas("PARC_Numero") & "/" & rsParcelas("Plano")
		
		'Preenche o andamento da cobrança
		'if Session("CodigoCliente") <> 0 then
		
		if Session("CodigoCliente") = 20 then
			BD.Execute "EXEC prc_CriaAndamentoCobranca @CTRA_ID = " & rsParcelas("CTRA_ID") & ", @ANCO_DataHora = '" & vHora & "', @STAC_ID = 23, @TDEV_ID = NULL, @ANCO_Descricao = 'Emissão do boleto " & rsParcelas("PARC_NumDocumento") & ", referente à parcela " & rsParcelas("PARC_Numero") & ". Venct. " & var_datavencimento & ", validade " & vValidadeBoleto  & ". Valor R$ "& var_valordocumento &"', @FUNC_ID = " & Session("FUNC_ID")
			PreencheLOG "EXEC prc_CriaAndamentoCobranca @CTRA_ID = " & rsParcelas("CTRA_ID") & ", @ANCO_DataHora = '" & vHora & "', @STAC_ID = 23, @TDEV_ID = NULL, @ANCO_Descricao = 'Emissão do boleto " & rsParcelas("PARC_NumDocumento") & ", referente à parcela " & rsParcelas("PARC_Numero") & ". Venct. " & var_datavencimento & ", validade " & vValidadeBoleto  & ". Valor R$ "& var_valordocumento &"', @FUNC_ID = " & Session("FUNC_ID"), vFilialDestino
		else
			BD.Execute "EXEC prc_CriaAndamentoCobranca @CTRA_ID = " & rsParcelas("CTRA_ID") & ", @ANCO_DataHora = '" & vHora & "', @STAC_ID = 23, @TDEV_ID = NULL, @ANCO_Descricao = 'Emissão do boleto " & rsParcelas("PARC_NumDocumento") & ", referente à parcela " & rsParcelas("PARC_Numero") & ". Venct. " & var_datavencimento & ", validade " & vValidadeBoleto  & ".', @FUNC_ID = " & Session("FUNC_ID")
			PreencheLOG "EXEC prc_CriaAndamentoCobranca @CTRA_ID = " & rsParcelas("CTRA_ID") & ", @ANCO_DataHora = '" & vHora & "', @STAC_ID = 23, @TDEV_ID = NULL, @ANCO_Descricao = 'Emissão do boleto " & rsParcelas("PARC_NumDocumento") & ", referente à parcela " & rsParcelas("PARC_Numero") & ". Venct. " & var_datavencimento & ", validade " & vValidadeBoleto  & ".', @FUNC_ID = " & Session("FUNC_ID"), vFilialDestino
		end if
		
	elseif vTipoBoleto = 4 then
		if vCont_id = 43 then
			var_contrato = rsParcelas("CTRA_Conta") & rsParcelas("CTRA_Numero") & " - " & rsBanco("CONT_Fantasia")
		else
			var_contrato = FormataContrato(rsParcelas("CTRA_Numero")) & " - " & rsBanco("CONT_Fantasia")
		end if
		if vCont_id = 3 then
			var_parcelaplano = "1/1"
		else
			var_parcelaplano = rsParcelas("TRAN_NumTitulo")
		end if
		
		vValorTaxaBoleto = rsParcelas("BOAV_TaxaBoleto")

		var_datavencimento = rsParcelas("BOAV_Vencimento") '"05/06/2002"
		var_valordocumento = FormatNumber(CDbl(vValorTaxaBoleto) + CDbl(rsParcelas("BOAV_Valor"))) '"70,00"

		if Session("CodigoCliente") = 13 and vCONT_ID = 11 then
			vTel = "(21)3212-3771"
		elseif Session("CodigoCliente") = 13 and (rsParcelas("CART_ID") = 38001 or rsParcelas("CART_ID") = 47001 or rsParcelas("CART_ID") = 42001 or rsParcelas("CART_ID") = 48001) then
			vTel = "(21)3212-3792"
		elseif Session("CodigoCliente") = 13 and vCONT_ID = 27 then
			vTel = "(21)3212-3780"
		elseif Session("CodigoCliente") = 13 and vCONT_ID = 14 then
			vTel = "(21)3212-3788"
		elseif Session("CodigoCliente") = 13 and (vCONT_ID = 16 or vCONT_ID = 17 or vCONT_ID = 61) then
			vTel = "(21)3212-3789"
		elseif Session("CodigoCliente") = 13 and vCONT_ID = 77 then
			vTel = "(21)3212-3793"
		elseif Session("CodigoCliente") = 13 and vCONT_ID = 80 then
			vTel = "(21)3212-3794"
		elseif Session("CodigoCliente") = 13 and vCONT_ID = 86 then
			vTel = "(21)3212-3796"
		elseif Session("CodigoCliente") = 13 and vCONT_ID = 48 then
			vTel = "(21)3212-3786"
		elseif Session("CodigoCliente") = 1 and vCONT_ID = 85 then
			vTel = "(71)3483-1250"
		elseif Session("CodigoCliente") = 23 and vCONT_ID = 58 then
			vTel = "(21)3984-2705"
		elseif Session("CodigoCliente") = 25 and vCONT_ID = 54 then
			vTel = "São Paulo (11) 2165-9423 - Demais localidades 0800 724 23 23"
		elseif Session("CodigoCliente") = 25 and (vCONT_ID = 100 or vCONT_ID = 104) then
			vTel = "SP e Gde. SP: (11) 2165-9436 - Outras Localidades:  0800-7241100"			
		end if

		if vTel <> "" then
			vTel = " Tel.: " & vTel
		end if
		if vFax <> "" then
			vFax = " Fax.: " & vFax
		end if

		vValidadeBoleto = rsParcelas("BOAV_Validade")
		if vCont_id = 3 then
			vQtdContratosHSBC = 0
			vContratos = rsParcelas("TRAN_NumTitulo") & " - " & rsBanco("CONT_Fantasia")
			var_contrato = vContratos
			var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & rsParcelas("BOAV_Validade") & ". <BR>Após " & rsParcelas("BOAV_Validade") & " entrar em contato com " & cons_cedente2 & " - " & vTel & " " & vFax & ". <BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & vContratos & " - Parcela/Plano: " & var_parcelaplano & " - CNPJ/CPF: " & FormataCPFCNPJ(rsParcelas("DEVE_CGCCPF")) & " <br>O envio desse boleto não inibe uma possível Ação de Débito efetuado pelo HSBC na conta corrente.</b>" 
		elseif vCont_id = 4 then
			var_contrato = Mid(var_contrato,1 ,2) & "**.****.****." & Mid(var_contrato,16 ,4)
			var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & rsParcelas("BOAV_Validade") & ". <BR>Após " & rsParcelas("BOAV_Validade") & " entrar em contato com " & cons_cedente2 & " - " & vTel & " " & vFax & ". <BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " - CNPJ/CPF: " & FormataCPFCNPJ(rsParcelas("DEVE_CGCCPF")) & " <br>O envio desse boleto não inibe uma possível Ação de Débito efetuado pelo HSBC na conta corrente.</b>" 
		elseif vCont_id = 58 then
			var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & rsParcelas("BOAV_Validade") & ". <BR>Após " & rsParcelas("BOAV_Validade") & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & ").<BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente a(s) parcela(s) "& var_parcelaplano & "  do contrato nº: " & var_contrato & "</b>. O não pagamento até a data de vencimento acarretará no cancelamento do acordo e a inclusão do seu nome no órgão de proteção ao crédito de sua região. " 
		elseif vCont_id = 80 then

		Set rsChequeInfo = Bd.Execute("SELECT DADO_Valor FROM Dados_Adicionais_do_Contrato WHERE CTRA_ID = " & vCTRA_ID & " AND TDAD_ID = 87")

		var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & rsParcelas("BOAV_Validade") & ". <BR>Após " & rsParcelas("BOAV_Validade") & " entrar em contato com " & cons_cedente2 & " - " & vTel & " " & vFax & ". <BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " </b>" 
		
		if not rsChequeInfo.EOF then
			var_instrucoes = var_instrucoes & " - " & rsChequeInfo("DADO_Valor") 
		end if 

		else

			if CDbl(vValorTaxaBoleto) > 0 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & rsParcelas("BOAV_Validade") & ". <BR>Após " & rsParcelas("BOAV_Validade") & " entrar em contato com " & cons_cedente2 & " - " & vTel & " " & vFax & ". <BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " </b>" 
			else
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & rsParcelas("BOAV_Validade") & ". <BR>Após " & rsParcelas("BOAV_Validade") & " entrar em contato com " & cons_cedente2 & " - " & vTel & " " & vFax & ". <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " </b>" 
			end if
		end if
		var_textoparcela = rsParcelas("TRAN_NumTitulo")

		'Preenche o andamento da cobrança
		'if Session("CodigoCliente") <> 0 then
		BD.Execute "EXEC prc_CriaAndamentoCobranca @CTRA_ID = " & rsParcelas("ctra_id") & ", @ANCO_DataHora = '" & vHora & "', @STAC_ID = 23, @TDEV_ID = NULL, @ANCO_Descricao = 'Emissão do boleto de título " & var_nossonumero & ", no valor de R$ " & var_valordocumento & ", com vencimento em " & var_datavencimento & ", referente ao título " & var_parcelaplano & "', @FUNC_ID = " & Session("FUNC_ID")
		PreencheLOG "EXEC prc_CriaAndamentoCobranca @CTRA_ID = " & rsParcelas("ctra_id") & ", @ANCO_DataHora = '" & vHora & "', @STAC_ID = 23, @TDEV_ID = NULL, @ANCO_Descricao = 'Emissão do boleto de título " & var_nossonumero & ", no valor de R$ " & var_valordocumento & ", com vencimento em " & var_datavencimento & ", referente ao título " & var_parcelaplano & "', @FUNC_ID = " & Session("FUNC_ID"), vFilialDestino
		'end if
				
		if rsParcelas("SCON_ID") <> 2 then
			vHora = DateAdd("s", 1, vHora)
			strSQL = "EXEC prc_CriaAndamentoCobranca @CTRA_ID = " & rsParcelas("ctra_id") & ", @ANCO_DataHora = '" & vHora & "', @STAC_ID = 13, @TDEV_ID = NULL, @ANCO_Descricao = 'Contrato direcionado para a fila de cobrança Pré-Acordo. Motivo: Boleto de título emitido.', @FUNC_ID = 0"
			BD.Execute strSQL
			PreencheLOG strSQL, vFilialDestino

			BD.Execute "prc_INS_HistoricoAgendamento_Forcado 3001, '" & now & "', " & rsParcelas("ctra_id") & ", " & Session("FILI_ID") & ", " & Session("FUNC_ID")
			PreencheLOG "prc_INS_HistoricoAgendamento_Forcado 3001, '" & now & "', " & rsParcelas("ctra_id") & ", " & Session("FILI_ID") & ", " & Session("FUNC_ID"), vFilialDestino
		end if

	elseif vTipoBoleto = 2 then
		
		if vCont_id = 43 then
			var_contrato = rsParcelas("CTRA_Conta") & rsParcelas("CTRA_Numero") & " - " & rsBanco("CONT_Fantasia")
		else
			var_contrato = FormataContrato(rsParcelas("CTRA_Numero")) & " - " & rsBanco("CONT_Fantasia")
		end if
		var_parcelaplano = "1 / " & Request.Form("txtPlano")

		var_datavencimento = Request.Form("txtVencimento") '"05/06/2002"
		var_valordocumento = FormatNumber(CDbl(vValorTaxaBoleto) + CDbl(Request.Form("txtValor"))) '"70,00"

		if Session("CodigoCliente") = 13 and vCONT_ID = 11 then
			vTel = "(21)3212-3771"
		elseif Session("CodigoCliente") = 13 and (rsParcelas("CART_ID") = 38001 or rsParcelas("CART_ID") = 47001 or rsParcelas("CART_ID") = 42001 or rsParcelas("CART_ID") = 48001) then
			vTel = "(21)3212-3792"
		elseif Session("CodigoCliente") = 13 and vCONT_ID = 27 then
			vTel = "(21)3212-3780"
		elseif Session("CodigoCliente") = 13 and vCONT_ID = 14 then
			vTel = "(21)3212-3788"
		elseif Session("CodigoCliente") = 13 and (vCONT_ID = 16 or vCONT_ID = 17 or vCONT_ID = 61) then
			vTel = "(21)3212-3789"
		elseif Session("CodigoCliente") = 13 and vCONT_ID = 77 then
			vTel = "(21)3212-3793"
		elseif Session("CodigoCliente") = 13 and vCONT_ID = 80 then
			vTel = "(21)3212-3794"
		elseif Session("CodigoCliente") = 13 and vCONT_ID = 86 then
			vTel = "(21)3212-3796"
		elseif Session("CodigoCliente") = 13 and vCONT_ID = 48 then
			vTel = "(21)3212-3786"
		elseif Session("CodigoCliente") = 1 and vCONT_ID = 85 then
			vTel = "(71)3483-1250"
		elseif Session("CodigoCliente") = 23 and vCONT_ID = 58 then
			vTel = "(21)3984-2705"
		elseif Session("CodigoCliente") = 25 and vCONT_ID = 54 then
			vTel = "São Paulo (11) 2165-9423 - Demais localidades 0800 724 23 23"
		elseif Session("CodigoCliente") = 25 and (vCONT_ID = 100 or vCONT_ID = 104) then
			vTel = "SP e Gde. SP: (11) 2165-9436 - Outras Localidades:  0800-7241100"
		end if

		if vTel <> "" then
			vTel = " Tel.: " & vTel
		end if
		if vFax <> "" then
			vFax = " Fax.: " & vFax
		end if

		vValidadeBoleto = DateAdd("d", Request.Form("txtValidade"), var_datavencimento)
		if vCont_id = 1 then
			if CDbl(vValorTaxaBoleto) > 0 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & DateAdd("d", Request.Form("txtValidade"), var_datavencimento) & ". <BR>Após " & DateAdd("d", Request.Form("txtValidade"), var_datavencimento) & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "). <BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " <br>O não pagamento até a data de vencimento acarretará no cancelamento do acordo e a inclusão do seu nome no órgão proteção ao crédito de sua região e no SERASA. Implicando na perda do desconto porventura concedido.</b>" 
			else
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & DateAdd("d", Request.Form("txtValidade"), var_datavencimento) & ". <BR>Após " & DateAdd("d", Request.Form("txtValidade"), var_datavencimento) & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "). <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " <br>O não pagamento até a data de vencimento acarretará no cancelamento do acordo e a inclusão do seu nome no órgão proteção ao crédito de sua região e no SERASA. Implicando na perda do desconto porventura concedido.</b>" 
			end if
		elseif vCont_id = 8 OR vCont_id = 9 then
			if CDbl(vValorTaxaBoleto) > 0 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & DateAdd("d", Request.Form("txtValidade"), var_datavencimento) & ". <BR>Após " & DateAdd("d", Request.Form("txtValidade"), var_datavencimento) & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "). <BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " <br>O não pagamento até a data de vencimento acarretará no cancelamento do acordo e a inclusão do seu nome no órgão proteção ao crédito de sua região.</b>" 
			else
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & DateAdd("d", Request.Form("txtValidade"), var_datavencimento) & ". <BR>Após " & DateAdd("d", Request.Form("txtValidade"), var_datavencimento) & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "). <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " <br>O não pagamento até a data de vencimento acarretará no cancelamento do acordo e a inclusão do seu nome no órgão proteção ao crédito de sua região.</b>" 
			end if
		elseif vCont_id = 2 then
			var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & DateAdd("d", Request.Form("txtValidade"), var_datavencimento) & ". <BR>Após o vencimento cobrar 2% de multa. <BR>Após " & DateAdd("d", Request.Form("txtValidade"), var_datavencimento) & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " </b>" 
		elseif vCont_id = 43 then
			var_instrucoes = "<B>Pagável em qualquer banco até o dia " & DateAdd("d", Request.Form("txtValidade"), var_datavencimento) & ". <BR>Após " & DateAdd("d", Request.Form("txtValidade"), var_datavencimento) & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & ".<br>CGC/CPF: " & FormataCPFCNPJ(rsParcelas("DEVE_CGCCPF")) & ".</b>" 
		elseif vCont_id = 11 then
			if CDbl(vValorTaxaBoleto) > 0 then
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & DateAdd("d", Request.Form("txtValidade"), var_datavencimento) & ". <BR>Após " & DateAdd("d", Request.Form("txtValidade"), var_datavencimento) & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido.  <BR>Valor do documento acrescido de " & FormatCurrency(vValorTaxaBoleto) & " referente a tarifa bancária. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " </b>" 
			else
				var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & DateAdd("d", Request.Form("txtValidade"), var_datavencimento) & ". <BR>Após " & DateAdd("d", Request.Form("txtValidade"), var_datavencimento) & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " </b>" 
			end if
		else
			var_instrucoes = "<BR> <B>Pagável em qualquer banco até o dia " & DateAdd("d", Request.Form("txtValidade"), var_datavencimento) & ". <BR>Após " & DateAdd("d", Request.Form("txtValidade"), var_datavencimento) & " pagar somente no Cedente (" & cons_cedente2 & " - " & vTel & " " & vFax & "), estando o acordo sujeito a cancelamento, implicando na perda do desconto porventura concedido. <BR>Referente ao contrato nº: " & var_contrato & " - Parcela/Plano: " & var_parcelaplano & " </b>" 
		end if
		var_textoparcela = "1 / " & Request.Form("txtPlano")
		'Preenche o andamento da cobrança
		'if Session("CodigoCliente") <> 0 then
		BD.Execute "EXEC prc_CriaAndamentoCobranca @CTRA_ID = " & Request.Form("ctra_id") & ", @ANCO_DataHora = '" & vHora & "', @STAC_ID = 23, @TDEV_ID = NULL, @ANCO_Descricao = 'Emissão do boleto avulso " & var_nossonumero & ", no valor de R$ " & var_valordocumento & ", com vencimento em " & var_datavencimento & ", referente à parcela/plano " & var_parcelaplano & "', @FUNC_ID = " & Session("FUNC_ID")
		PreencheLOG "EXEC prc_CriaAndamentoCobranca @CTRA_ID = " & Request.Form("ctra_id") & ", @ANCO_DataHora = '" & vHora & "', @STAC_ID = 23, @TDEV_ID = NULL, @ANCO_Descricao = 'Emissão do boleto avulso " & var_nossonumero & ", no valor de R$ " & var_valordocumento & ", com vencimento em " & var_datavencimento & ", referente à parcela/plano " & var_parcelaplano & "', @FUNC_ID = " & Session("FUNC_ID"), vFilialDestino
		'end if
		
		if rsParcelas("SCON_ID") <> 2 then
			vHora = DateAdd("s", 1, vHora)
			strSQL = "EXEC prc_CriaAndamentoCobranca @CTRA_ID = " & Request.Form("ctra_id") & ", @ANCO_DataHora = '" & vHora & "', @STAC_ID = 13, @TDEV_ID = NULL, @ANCO_Descricao = 'Contrato direcionado para a fila de cobrança Pré-Acordo. Motivo: Boleto avulso emitido.', @FUNC_ID = 0"
			BD.Execute strSQL
			PreencheLOG strSQL, vFilialDestino

			BD.Execute "prc_INS_HistoricoAgendamento_Forcado 3001, '" & now & "', " & Request.Form("ctra_id") & ", " & Session("FILI_ID") & ", " & Session("FUNC_ID")
			PreencheLOG "prc_INS_HistoricoAgendamento_Forcado 3001, '" & now & "', " & Request.Form("ctra_id") & ", " & Session("FILI_ID") & ", " & Session("FUNC_ID"), vFilialDestino
		end if
	end if

	if Session("CodigoCliente") = 8 then
		var_instrucoes = var_instrucoes & "<BR><b>Somente efetuar pagamento na Rede Bancária, não deve ser pago em agentes arrecadadores.</b>"
	elseif Session("CodigoCliente") = 11 then
		var_instrucoes = var_instrucoes & "<BR><b>Prezado Cliente, não efetuar o pagamento na Fininvest.</b>"
	end if

	if Request.QueryString("recuperador") <> "" then
		vValid = rsParcelas("BOAV_Validade")
	elseif Request.QueryString("automatico") <> "" then
		vValid = var_datavencimento
	elseif Request.QueryString("atraso") <> "" then
		vValid = rsParcelas("BOAV_Validade")
	else
		vValid = Request.Form("txtValidade" & rsParcelas("PARC_ID"))
		if vValid = "" then
			vValid = rsParcelas("BOAV_Validade")
		else
			vValid = DateAdd("d", Request.Form("txtValidade" & rsParcelas("PARC_ID")), var_datavencimento)
		end if
	end if
	if vValid = "0" then
		vValid = DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento)
	end if
	
	if vTpBoleto = "" then
		vTpBoleto = "A"
	end if
	if vTpBoleto = "P" then
		BD.Execute "UPDATE Boletos_Gerados SET EDEV_ID = " & rsParcelas("EDEV_ID") & ", BOGE_Valor = " & Replace(Replace(var_valordocumento - vValorTaxaBoleto, ".", ""), ",", ".") & ", BOGE_Vencimento = '" & var_datavencimento & "', BOGE_Validade = '" & vValid & "', BOGE_TaxaBoleto = " & Replace(Replace(vValorTaxaBoleto, ".", ""), ",", ".") & " WHERE PARC_ID = " & rsParcelas("PARC_ID") & " AND CTRA_ID = " & rsParcelas("CTRA_ID") & " AND BOGE_TipoBoleto = 'P' AND BOGE_Numero = " & rsParcelas("PARC_NumDocumento")
		PreencheLOG "UPDATE Boletos_Gerados SET EDEV_ID = " & rsParcelas("EDEV_ID") & ", BOGE_Valor = " & Replace(Replace(var_valordocumento - vValorTaxaBoleto, ".", ""), ",", ".") & ", BOGE_Vencimento = '" & var_datavencimento & "', BOGE_Validade = '" & vValid & "', BOGE_TaxaBoleto = " & Replace(Replace(vValorTaxaBoleto, ".", ""), ",", ".") & " WHERE PARC_ID = " & rsParcelas("PARC_ID") & " AND CTRA_ID = " & rsParcelas("CTRA_ID") & " AND BOGE_TipoBoleto = 'P' AND BOGE_Numero = " & rsParcelas("PARC_NumDocumento"), vFilialDestino
	else
		BD.Execute "UPDATE Boletos_Gerados SET EDEV_ID = " & rsParcelas("EDEV_ID") & ", BOGE_Valor = " & Replace(Replace(var_valordocumento - vValorTaxaBoleto, ".", ""), ",", ".") & ", BOGE_Vencimento = '" & var_datavencimento & "', BOGE_Validade = '" & vValid & "', BOGE_TaxaBoleto = " & Replace(Replace(vValorTaxaBoleto, ".", ""), ",", ".") & " WHERE PARC_ID IS NULL AND CTRA_ID = " & rsParcelas("CTRA_ID") & " AND BOGE_TipoBoleto = '" & vTpBoleto & "' AND BOGE_Numero = " & rsParcelas("PARC_NumDocumento")
		PreencheLOG "UPDATE Boletos_Gerados SET EDEV_ID = " & rsParcelas("EDEV_ID") & ", BOGE_Valor = " & Replace(Replace(var_valordocumento - vValorTaxaBoleto, ".", ""), ",", ".") & ", BOGE_Vencimento = '" & var_datavencimento & "', BOGE_Validade = '" & vValid & "', BOGE_TaxaBoleto = " & Replace(Replace(vValorTaxaBoleto, ".", ""), ",", ".") & " WHERE PARC_ID IS NULL AND CTRA_ID = " & rsParcelas("CTRA_ID") & " AND BOGE_TipoBoleto = '" & vTpBoleto & "' AND BOGE_Numero = " & rsParcelas("PARC_NumDocumento"), vFilialDestino
	end if

	if vCONT_ID = 74 then
		var_instrucoes = "<br><b>Não receber após o vencimento. Não receber em cheque. Débito referente ao Banco Sudameris Brasil S.A.<br>Ciente: ""Reconheço e pagarei a dívida acima nas condições aqui oferecidas. Fico ciente de que, caso não venha a cumprir com os valores e prazos fixados, tornar-se-ão sem efeito os descontos propostos, não se tratando de novação"".<br>Código negociador: " & vContIdentArq & "<b>"
	elseif vCONT_ID = 75 then
		var_instrucoes = "<br><b>Não receber após o vencimento. Não receber em cheque. Débito referente ao Banco ABN AMRO Real S.A.<br>Ciente: ""Reconheço e pagarei a dívida acima nas condições aqui oferecidas. Fico ciente de que, caso não venha a cumprir com os valores e prazos fixados, tornar-se-ão sem efeito os descontos propostos, não se tratando de novação"".<br>Código negociador: " & vContIdentArq & "<b>"
	elseif vCONT_ID = 77 then
		if CDate(var_datavencimento) > CDate("23/12/2005") then
			var_instrucoes = "<br>Aproveite esta oportunidade para quitar seu débito. Tenha o seu Guanabara Card liberado para compras  após quitação e  análise de crédito,<font size=2> <B>sem consulta ao SPC/SERASA</B></font>. Para isso, sete dias úteis após quitação, entre em contato com tel.: 2157-5858. Sr caixa, não receber após vencimento. Após vencimento, pagável somente na "& cons_cedente &" &nbsp; (Endereço: "& vEndFilial &" - "& vTel &")."
		else
			var_instrucoes = "<br>Seu Guanabara Card será liberado para compras após quitação e processamento total do pagamento. <br> Para isso, após quitação, entre em contato pelo "& vTel &" <br> ATENÇÃO: ESTA PROMOÇÃO É VÁLIDA SOMENTE ATÉ 23/12/2005. APROVEITE! <br> Sr Caixa, não receber após vencimento <br> Após vencimento pagável somente na " & vNomeEmpresa & " - " & vEndFilial
		end if
	end if
	if rsBanco("BANC_ID_Boleto") = "008" or rsBanco("BANC_ID_Boleto") = "353" then
		SDIG=""
		CDIG=""
		LDIG=""
		NOSSONUMERO=""

		'********************************
		' CONSTANTES
		'********************************

		if rsBanco("BANC_ID_Boleto") = "008" then
			cons_banco = "008"
			cons_dvbanco = "6"
		else 
			cons_banco = "353"
			cons_dvbanco = "0"
		end if
		cons_carteira = "CSR"

		var_numerodoc = var_nossonumero '"008171001A"
		var_localpagamento = "PAGÁVEL EM QUALQUER BANCO ATÉ O VENCIMENTO"

		'********************************
		' INICIO DO CÁLCULO
		'********************************

		dvnossonumero = CALCDIG11_Santander(var_nossonumero,9,0) 'calcdig10(var_nossonumero)

		'dvconta = CALCDIG10(cons_conta) 'calcdig10(cons_conta)
		dvagencia = CALCDIG10(cons_agencia) 'calcdig10(cons_agencia)

		valordia = date()
		'var_data = Day(valordia) & "/" & Month(valordia) & "/" & YEAR(valordia)
		var_data = valordia

		valorvalor1=var_valordocumento
		valorvalor2=replace(valorvalor1,",","")
		valorvalor2=replace(valorvalor2,".","")
		valorvalor3=len(valorvalor2)
		valorvalor4=10-valorvalor3
		var_valor= String(""&valorvalor4&"","0") & (""&valorvalor2&"")
		if valorvalor1=0 then
			var_valor=""
		end if

		'var_fatorvencimento=fatorvencimento(""& var_datavencimento &"")
		if vCont_id = 11 then
			var_fatorvencimento=fatorvencimento(""& var_datavencimento &"")
		else
			if Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")) = "" then
				var_fatorvencimento=fatorvencimento(""& vValid &"")
			else
				var_fatorvencimento=fatorvencimento(""& DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) &"")
			end if
		end if
		if var_fatorvencimento="0000" then
			var_datavencimento="Contra Apresentação"
		end if

		var_codigobarras=codbar_santander(""&cons_banco&"",""&cons_moeda&"",""&var_fatorvencimento&"",""&var_valor&"",""&cons_codcliente&"",""&var_nossonumero&dvnossonumero&"")
		var_linhadigitavel=linhadigitavel_santander(""&var_codigobarras&"")

		var_textocodigocedente = cons_agencia & "-" & dvagencia & " " & cons_codcliente
		var_textonossonumero = var_nossonumero & "-" & dvnossonumero
		var_textoaceite = " "
		
		var_colunas = 1
		var_tamcoluna = 170
		var_textoespecie = cons_especie
	elseif rsBanco("BANC_ID_Boleto") = "033" then
		SDIG=""
		CDIG=""
		LDIG=""
		NOSSONUMERO=""

		'********************************
		' CONSTANTES
		'********************************

		cons_banco = "033"
		cons_dvbanco = "7"
		cons_carteira = "COBR." 'cons_carteira '"COB"

		var_numerodoc = var_nossonumero
		var_localpagamento = "ATÉ O VENCIMENTO PAGÁVEL EM QUALQUER BANCO"

		'********************************
		' INICIO DO CÁLCULO
		'********************************

		dvnossonumero = CALCNUMB2(Trim(cons_agencia) & Trim(var_nossonumero))
		'dvagconta = calcdig10(cons_agencia&cons_conta)


		valordia = date()
		'var_data = Day(valordia) & "/" & Month(valordia) & "/" & YEAR(valordia)
		var_data = valordia

		valorvalor1 = var_valordocumento
		valorvalor2 = replace(valorvalor1,",","")
		valorvalor2 = replace(valorvalor2,".","")
		valorvalor3 = len(valorvalor2)
		valorvalor4 = 10-valorvalor3
		var_valor = String(""&valorvalor4&"","0") & (""&valorvalor2&"")
		if valorvalor1 = 0 then
			var_valor = ""
		end if

		var_fatorvencimento = fatorvencimento(""& vValidadeBoleto &"")
		'var_fatorvencimento=fatorvencimento(""& var_datavencimento &"")
		'if Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")) = "" then
		'	var_fatorvencimento=fatorvencimento(""& vValid &"")
		'else
		'	var_fatorvencimento=fatorvencimento(""& DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) &"")
		'end if
		if var_fatorvencimento="0000" then
			var_datavencimento="Contra Apresentação"
		end if

		var_codigobarras=codbar_banespa(""&cons_banco&"",""&cons_moeda&"",""&var_fatorvencimento&"",""&var_valor&"",""&cons_agencia & cons_conta&"",""&var_nossonumero&"")
		var_linhadigitavel=linhadigitavel_banespa(""&var_codigobarras&"")
		
		var_textocodigocedente = cons_agencia & "&nbsp;" & Mid(cons_conta, 1, 2) & "&nbsp;" & Mid(cons_conta, 3, 5) & "-" & Right(cons_conta, 1)
		var_textonossonumero = cons_agencia & "&nbsp;" & var_nossonumero & "&nbsp;" & dvnossonumero
		var_textoaceite = "N"
		
		var_colunas = 1
		var_tamcoluna = 170
		var_textoespecie = cons_especie
	elseif rsBanco("BANC_ID_Boleto") = "104" then
		SDIG=""
		CDIG=""
		LDIG=""
		NOSSONUMERO=""

		'********************************
		' CONSTANTES
		'********************************

		cons_banco = "104"
		cons_dvbanco = "0"
		cons_carteira = rsBanco("CONT_CarteiraBoleto") '"COB"

		var_numerodoc = var_nossonumero
		var_localpagamento = "ATÉ O VENCIMENTO PAGÁVEL EM QUALQUER BANCO"

		'********************************
		' INICIO DO CÁLCULO
		'********************************

		dvnossonumero = CALCDIG11("8" & Trim(var_nossonumero),9,0)
		dvagconta = CALCDIG11(cons_agencia & "870000" & cons_codcliente,9,0)

		valordia = date()
		'var_data = Day(valordia) & "/" & Month(valordia) & "/" & YEAR(valordia)
		var_data = valordia

		valorvalor1=var_valordocumento
		valorvalor2=replace(valorvalor1,",","")
		valorvalor2=replace(valorvalor2,".","")
		valorvalor3=len(valorvalor2)
		valorvalor4=10-valorvalor3
		var_valor= String(""&valorvalor4&"","0") & (""&valorvalor2&"")
		if valorvalor1=0 then
			var_valor=""
		end if

		'var_fatorvencimento=fatorvencimento(""& var_datavencimento &"")
		'if Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")) = "" then
			var_fatorvencimento=fatorvencimento(""& vValid &"")
		'else
		'	var_fatorvencimento=fatorvencimento(""& DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) &"")
		'end if
		if var_fatorvencimento="0000" then
			var_datavencimento="Contra Apresentação"
		end if

		var_codigobarras=codbar_caixa(""&cons_banco&"",""&cons_moeda&"",""&var_fatorvencimento&"",""&var_valor&"",""&cons_codcliente&""&"",""&cons_agencia&"",""&var_nossonumero&"")
		var_linhadigitavel=linhadigitavel_caixa(""&var_codigobarras&"")
		
		var_textocodigocedente = cons_agencia & ".870.000" & cons_codcliente & "-" & dvagconta
		var_textonossonumero = "8" & var_nossonumero & "-" & dvnossonumero
		var_textoaceite = "N"
		
		var_colunas = 1
		var_tamcoluna = 170
		var_textoespecie = cons_especie

	elseif rsBanco("BANC_ID_Boleto") = 237 then
		SDIG=""
		CDIG=""
		LDIG=""
		NOSSONUMERO=""

		'********************************
		' CONSTANTES
		'********************************

		cons_banco = "237"
		cons_dvbanco = "2"
		cons_carteira = cons_carteira '"175"

		var_numerodoc = var_nossonumero '"008171001A"
		var_localpagamento = "Pagável Preferencialmente em qualquer Agência Bradesco"

		'********************************
		' INICIO DO CÁLCULO
		'********************************

		dvnossonumero = CALCDIG11_Bradesco(cons_carteira & var_nossonumero,7,0) 'calcdig10(var_nossonumero)

		dvconta = CALCDIG11_Bradesco(cons_conta,7,0) 'calcdig10(cons_conta)
		dvagencia = CALCDIG11_Bradesco(cons_agencia,7,0) 'calcdig10(cons_agencia)

		valordia = date()
		'var_data = Day(valordia) & "/" & Month(valordia) & "/" & YEAR(valordia)
		var_data = valordia

		valorvalor1=var_valordocumento
		valorvalor2=replace(valorvalor1,",","")
		valorvalor2=replace(valorvalor2,".","")
		valorvalor3=len(valorvalor2)
		valorvalor4=10-valorvalor3
		var_valor= String(""&valorvalor4&"","0") & (""&valorvalor2&"")
		if valorvalor1=0 then
			var_valor=""
		end if

		'var_fatorvencimento=fatorvencimento(""& var_datavencimento &"")
		if vCont_id = 11 then
			var_fatorvencimento=fatorvencimento(""& var_datavencimento &"")
		else
			if Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")) = "" then
				var_fatorvencimento=fatorvencimento(""& vValid &"")
			else
				var_fatorvencimento=fatorvencimento(""& DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) &"")
			end if
		end if
		if var_fatorvencimento="0000" then
			var_datavencimento="Contra Apresentação"
		end if

		var_codigobarras=codbar_bradesco(""&cons_banco&"",""&cons_moeda&"",""&var_fatorvencimento&"",""&var_valor&"",""&cons_agencia&"",""&cons_carteira&"",""&var_nossonumero&"",""&cons_conta&"")
		var_linhadigitavel=linhadigitavel_bradesco(""&var_codigobarras&"")

		if cons_conta = "0073760" then
			dvconta = 8
		end if
		var_textocodigocedente = cons_agencia & "-" & dvagencia & "&nbsp;/&nbsp;" & cons_conta & "-" & dvconta
		var_textonossonumero = cons_carteira & "&nbsp;/&nbsp;" & var_nossonumero & "-" & dvnossonumero
		var_textoaceite = " "
		
		var_colunas = 1
		var_tamcoluna = 170
		var_textoespecie = cons_especie
	elseif rsBanco("BANC_ID_Boleto") = 341 then
		SDIG=""
		CDIG=""
		LDIG=""
		NOSSONUMERO=""

		'********************************
		' CONSTANTES
		'********************************

		cons_banco = "341"
		cons_dvbanco = "7"
		cons_carteira = cons_carteira '"175"

		var_numerodoc = var_nossonumero '"008171001A"
		var_localpagamento = "ATÉ O VENCIMENTO PAGÁVEL EM QUALQUER BANCO"

		'********************************
		' INICIO DO CÁLCULO
		'********************************

		dvnossonumero = calcdig10(cons_agencia&cons_conta&cons_carteira&var_nossonumero)
		dvagconta = calcdig10(cons_agencia&cons_conta)


		valordia = date()
		'var_data = Day(valordia) & "/" & Month(valordia) & "/" & YEAR(valordia)
		var_data = valordia

		valorvalor1=var_valordocumento
		valorvalor2=replace(valorvalor1,",","")
		valorvalor2=replace(valorvalor2,".","")
		valorvalor3=len(valorvalor2)
		valorvalor4=10-valorvalor3
		var_valor= String(""&valorvalor4&"","0") & (""&valorvalor2&"")
		if valorvalor1=0 then
			var_valor=""
		end if

		'var_fatorvencimento=fatorvencimento(""& var_datavencimento &"")
		if Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")) = "" then
			var_fatorvencimento=fatorvencimento(""& vValid &"")
		else
			var_fatorvencimento=fatorvencimento(""& DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) &"")
		end if
		if var_fatorvencimento="0000" then
			var_datavencimento="Contra Apresentação"
		end if

		var_codigobarras=codbar_itau(""&cons_banco&"",""&cons_moeda&"",""&var_fatorvencimento&"",""&var_valor&"",""&cons_carteira&"",""&var_nossonumero&"",""&dvnossonumero&"",""&cons_agencia&"",""&cons_conta&"",""&dvagconta&"")
		var_linhadigitavel=linhadigitavel_itau(""&var_codigobarras&"")

		var_textocodigocedente = cons_agencia & "&nbsp;/&nbsp;" & cons_conta & "-" & dvagconta
		var_textonossonumero = cons_carteira & "&nbsp;/&nbsp;" & var_nossonumero & "-" & dvnossonumero
		var_textoaceite = "N"
		
		var_colunas = 1
		var_tamcoluna = 170
		var_textoespecie = cons_especie
	elseif rsBanco("BANC_ID_Boleto") = 347 or rsBanco("BANC_ID_Boleto") = 356 then
		SDIG=""
		CDIG=""
		LDIG=""
		NOSSONUMERO=""

		'********************************
		' CONSTANTES
		'********************************

		if rsBanco("BANC_ID_Boleto") = 347 then
			cons_banco = "347"
			cons_dvbanco = "6"
		else
			cons_banco = "356"
			cons_dvbanco = "5"
		end if
		cons_carteira = cons_carteira '"175"

		var_numerodoc = var_nossonumero '"008171001A"
		var_localpagamento = "PAGÁVEL EM QUALQUER BANCO ATÉ O VENCIMENTO"

		'********************************
		' INICIO DO CÁLCULO
		'********************************

		dvnossonumero = calcdig10(var_nossonumero&cons_agencia&cons_conta&cons_carteira)
		'dvagconta = calcdig10(cons_agencia&cons_conta)


		valordia = date()
		'var_data = Day(valordia) & "/" & Month(valordia) & "/" & YEAR(valordia)
		var_data = valordia

		valorvalor1=var_valordocumento
		valorvalor2=replace(valorvalor1,",","")
		valorvalor2=replace(valorvalor2,".","")
		valorvalor3=len(valorvalor2)
		valorvalor4=10-valorvalor3
		var_valor= String(""&valorvalor4&"","0") & (""&valorvalor2&"")
		if valorvalor1=0 then
			var_valor=""
		end if

		'var_fatorvencimento=fatorvencimento(""& var_datavencimento &"")
		if Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")) = "" then
			var_fatorvencimento=fatorvencimento(""& vValid &"")
		else
			var_fatorvencimento=fatorvencimento(""& DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) &"")
		end if
		if var_fatorvencimento="0000" then
			var_datavencimento="Contra Apresentação"
		end if

		var_codigobarras=codbar_real(""&cons_banco&"",""&cons_moeda&"",""&cons_agencia&"",""&cons_conta&"",""&dvnossonumero&"",""&var_nossonumero&"",""&var_fatorvencimento&"",""&var_valor&"")
		var_linhadigitavel=linhadigitavel_real(""&var_codigobarras&"")

		var_textocodigocedente = cons_agencia & "/" & cons_conta & "/" & dvnossonumero
		var_textonossonumero = var_nossonumero 
		var_textoaceite = "A"
		
		var_colunas = 1
		var_tamcoluna = 170
		var_textoespecie = cons_especie
	elseif rsBanco("BANC_ID_Boleto") = 399 then
		'HSBC
		'********************************
		' CONSTANTES
		'********************************

		cons_banco = "399"
		cons_dvbanco = "9"
		cons_carteira = cons_carteira '"CNR"

		var_numerodoc = var_nossonumero
		var_localpagamento = "PAGAR PREFERENCIALMENTE EM AGÊNCIA DO HSBC"


		'********************************
		' INICIO DO CÁLCULO
		'********************************

		dvnossonumero = CALCDIG11_HSBC(var_numerodoc,9,0)
		dv2nossonumero = "4"
		dv3nossonumero = CALCDIG11_HSBC(CDbl(var_numerodoc & dvnossonumero & dv2nossonumero) + CLng(cons_codcliente) + (CLng(Day(var_datavencimento) & Right("0" & Month(var_datavencimento),2) & Right(Year(var_datavencimento),2))),9,0)
		
		dvagconta=calcdig10(cons_agencia&cons_conta)


		valordia=date()
		'var_data = Day(valordia) & "/" & Month(valordia) & "/" & YEAR(valordia)
		var_data = ""

		valorvalor1=var_valordocumento
		valorvalor2=replace(valorvalor1,",","")
		valorvalor2=replace(valorvalor2,".","")
		valorvalor3=len(valorvalor2)
		valorvalor4=10-valorvalor3
		var_valor= String(""&valorvalor4&"","0") & (""&valorvalor2&"")
		if valorvalor1=0 then
			var_valor=""
		end if

		'var_fatorvencimento=fatorvencimento(""& var_datavencimento &"")
		if Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")) = "" then
			var_fatorvencimento=fatorvencimento(""& vValid &"")
		else
			var_fatorvencimento=fatorvencimento(""& DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) &"")
		end if
		if var_fatorvencimento="0000" then
			var_datavencimento="Contra Apresentação"
		end if
		
		var_codigobarras=codbar_hsbc(""&cons_banco&"",""&cons_moeda&"",""&var_fatorvencimento&"",""&var_valor&"",""&cons_codcliente&"",""&Right("0000000000000" & var_nossonumero, 13)&"",""&var_datavencimento&"")
		var_linhadigitavel=linhadigitavel_hsbc(""&var_codigobarras&"")

		var_textocodigocedente = cons_codcliente
		var_textonossonumero = var_nossonumero & dvnossonumero & dv2nossonumero & dv3nossonumero
		var_textoaceite = ""

		var_colunas = 2
		var_tamcoluna = 120
		var_textoespecie = cons_moeda & " - " & cons_especie
	elseif rsBanco("BANC_ID_Boleto") = 409 then
		'********************************
		' CONSTANTES
		'********************************

		cons_banco = "409"
		cons_dvbanco = "0"
		cons_carteira = cons_carteira '"175"

		var_numerodoc = var_nossonumero '"008171001A"
		var_localpagamento = "ATÉ O VENCIMENTO PAGÁVEL EM QUALQUER BANCO"


		'********************************
		' INICIO DO CÁLCULO
		'********************************

		dvnossonumero=SUPERDIGITO(var_numerodoc)
		dvagconta=calcdig10(cons_agencia&cons_conta)


		valordia=date()
		'var_data = Day(valordia) & "/" & Month(valordia) & "/" & YEAR(valordia)
		var_data = valordia

		valorvalor1=var_valordocumento
		valorvalor2=replace(valorvalor1,",","")
		valorvalor2=replace(valorvalor2,".","")
		valorvalor3=len(valorvalor2)
		valorvalor4=10-valorvalor3
		var_valor= String(""&valorvalor4&"","0") & (""&valorvalor2&"")
		if valorvalor1=0 then
			var_valor=""
		end if

		'var_fatorvencimento=fatorvencimento(""& var_datavencimento &"")
		if Request.Form("txtValidade") <> "" then
			var_fatorvencimento=fatorvencimento(""& vValid &"")
		else
			var_fatorvencimento=fatorvencimento(""& DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) &"")
		end if
		if var_fatorvencimento="0000" then
			var_datavencimento="Contra Apresentação"
		end if
		
		var_codigobarras=codbar_unibanco(""&cons_banco&"",""&cons_moeda&"",""&var_fatorvencimento&"",""&var_valor&"",""&cons_codcliente&"",""&var_nossonumero&"")
		var_linhadigitavel=linhadigitavel_unibanco(""&var_codigobarras&"")

'		if vCONT_ID = 77 then
			var_textocodigocedente = cons_agencia & "&nbsp;/&nbsp;" & CLng(cons_conta) & "-" & dvagconta
			var_textonossonumero = var_nossonumero & "/" & dvnossonumero
'		else
'			var_textocodigocedente = cons_agencia & "&nbsp;/&nbsp;" & CLng(cons_conta) & "-" & dvagconta
			'var_textonossonumero = "1/" & Right(var_nossonumero,11) & "/" & dvnossonumero
'			var_textonossonumero = var_nossonumero & "/" & dvnossonumero
'		end if
		var_textoaceite = "N"

		var_colunas = 1
		var_tamcoluna = 170
		var_textoespecie = cons_especie
	elseif rsBanco("BANC_ID_Boleto") = 623 then
		'PANAMERICANO
		'********************************
		' CONSTANTES
		'********************************

		cons_banco = "623"
		cons_dvbanco = "8"
		cons_carteira = cons_carteira '"CNR"

		var_numerodoc = var_nossonumero
		var_localpagamento = "PAGÁVEL EM QUALQUER BANCO ATÉ O VENCIMENTO"
		
		cons_cedente = "BANCO PANAMERICANO - " & cons_codcliente

		'********************************
		' INICIO DO CÁLCULO
		'********************************

		dvnossonumero = CALCDIG10(cons_agencia & cons_carteira & var_numerodoc)
		
		dvconta = calcdig11(cons_agencia&cons_conta,9,0) 'calcdig10(cons_conta)
		dvagencia = calcdig11(cons_agencia,9,0) 'calcdig10(cons_agencia)

		valordia=date()
		'var_data = Day(valordia) & "/" & Month(valordia) & "/" & YEAR(valordia)
		var_data = ""

		valorvalor1=var_valordocumento
		valorvalor2=replace(valorvalor1,",","")
		valorvalor2=replace(valorvalor2,".","")
		valorvalor3=len(valorvalor2)
		valorvalor4=10-valorvalor3
		var_valor= String(""&valorvalor4&"","0") & (""&valorvalor2&"")
		if valorvalor1=0 then
			var_valor=""
		end if

		'var_fatorvencimento=fatorvencimento(""& var_datavencimento &"")
		if Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")) = "" then
			var_fatorvencimento=fatorvencimento(""& vValid &"")
		else
			var_fatorvencimento=fatorvencimento(""& DateAdd("d", Request.Form("txtPrazoPagamento" & rsParcelas("PARC_ID")), var_datavencimento) &"")
		end if
		if var_fatorvencimento="0000" then
			var_datavencimento="Contra Apresentação"
		end if
		
		var_codigobarras=codbar_panamericano(""&cons_banco&"",""&cons_moeda&"",""&var_fatorvencimento&"",""&var_valor&"",""&cons_agencia&"",""&cons_carteira&"",""&cons_operacao&"",""&var_numerodoc&dvnossonumero&"")
		var_linhadigitavel=linhadigitavel_panamericano(""&var_codigobarras&"")

		var_textocodigocedente = cons_agencia & "-" & dvagencia & "/" & cons_conta & "-" & dvconta
		'var_textonossonumero = cons_agencia & dvagencia & "/" & cons_carteira & "/" & var_numerodoc & "-" & dvnossonumero
		var_textonossonumero = var_numerodoc & "-" & dvnossonumero
		var_textoaceite = ""

		var_colunas = 1
		var_tamcoluna = 170
		var_textoespecie = cons_moeda & " - " & cons_especie
	end if
	
	'Response.Write var_codigobarras & "<BR>"
	'Response.Write var_linhadigitavel & "<BR>"
	
	if vCONT_ID = 40 then
		cons_cedente = "OI – TNL PCS S/A"
	elseif vCONT_ID = 48 and rsBanco("BANC_ID_Boleto") = 104 then
		cons_cedente = "IBI ADM PROMOTORA LTDA"
	elseif vCONT_ID = 74 then
		cons_cedente = "Banco Sudameris do Brasil S/A"
	elseif vCONT_ID = 75 then
		cons_cedente = "Banco ABN Amro Real S/A"
	elseif vCONT_ID = 77 then
		cons_cedente = "Casas Guanabara Comestiveis Ltda"
	end if
	
	 
	if Session("CodigoCliente") = 26 then
		DescBoletoParcela rsParcelas("CTRA_Numero"), vNomeEmpresa, var_parcelaplano, var_datavencimento, vTel, var_valordocumento
	end if 
	
	if (vCONT_ID = 74 or vCONT_ID = 75) and (Request.QueryString("automatico") <> "" or Request.QueryString("atraso") <> "" or vTipoBoleto = 1) then
	%>
	<TABLE WIDTH="640" CELLSPACING=0 CELLPADDING=0 BORDER=0 ID="Table19">
  		<TR>
			<TD class=cp VALIGN=BOTTOM WIDTH=225><IMG SRC="../images/logobanco_<% = rsBanco("BANC_ID_Boleto")%>.gif"></TD>
			<TD ALIGN=RIGHT VALIGN=BOTTOM><FONT class=ld><B>RECIBO DO SACADO</B></FONT></TD>
		</TR>
	</TABLE>
	<TABLE WIDTH="640" BORDER=1 CELLSPACING=0 CELLPADDING=1 ID="Table20">
		<TR>
   			<TD COLSPAN=2><FONT class=ct>Sacado</FONT><BR><FONT class=cp>&nbsp;<%=var_sacado%></FONT></TD>
   			<TD width=15%><FONT class=ct>Número Documento</FONT><BR><FONT class=cn>&nbsp;<%=var_numerodoc%></FONT></TD>
			<TD width=20% bgcolor="#CCCCCC">
			<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0 ID="Table21">
					<TR><TD align=left><FONT class=ct>Vencimento</FONT></TD></TR>
						<TR><TD align=center><FONT class=cp><%=var_datavencimento%></FONT></TD></TR>
				</TABLE>
			</TD>
			</TD>
		</TR>
	</TABLE>
	<%
	if vCONT_ID = 74 then
	%>
    <table width="640" cellpadding="3" cellspacing="0" align="center" border=0 bordercolor=black height=460 ID="Table1">
	<%
	else
	%>
    <table width="640" cellpadding="3" cellspacing="0" align="center" border=0 bordercolor=black height=470 ID="Table2">
	<%
	end if
	%>
     <tr>  
      	<td valign=top>
			<br>
			<table width="620" cellpadding="0" cellspacing="0" align="center" border=0 ID="Table26">
				<TR>
					<td class=texto1>Ciente: "Reconheço e pagarei a dívida acima nas condições aqui oferecidas. Fico ciente de que, caso não venha a cumprir com os valores e prazos fixados, ternar-se-ão sem efeito os descontos propostos, não se tratando de novação".</td>
				</tr>
				<TR>
					<td class=texto1>Esta proposta de acordo compreende apenas o(s) debito(s) referente(s) a(s) parcela(s) do(s) contrato(s) abaixo relacionado(s).</td>
				</tr>
				<TR>
					<td class=texto1>A exclusão de SPC/Serasa fica condicionada à inexistência de outro débito vencido e não descrito nesta proposta acordo.</td>
				</tr>
				<TR>
					<td class=texto1>O pagamento do boleto não implica a reativação de conta corrente e limites de crédito, cuja análise estará sujeita ao <% if vCONT_ID = 74 then Response.Write "Banco Sudameris Brasil S.A." else Response.Write "Banco ABN Amro Real S.A."%>.</td>
				</tr>
				<TR>
					<td class=texto1>Caso a situação já tenha sido regularizada, favor desconsiderar este documento.</td>
				</tr>
			</table>
			<br>
			<table width="620" cellpadding="0" cellspacing="0" align="center" border=0 ID="Table25">
				<TR>
					<td class=texto1>Valor total do acordo: <% = FormatCurrency(rsParcelas("ValorAcordo"))%></td>
				</tr>
				<TR>
					<td class=texto1>Quantidade de parcelas do acordo: <% = rsParcelas("Plano")%> parcela(s)</td>
				</tr>
				<TR>
					
					<td class=texto1>Número da parcela do acordo: <%=rsParcelas("PARC_Numero")%></td>
				</tr>
				<TR>
					<td class=texto1>Valor da parcela: <% = FormatCurrency(var_valordocumento)%></td>
				</tr>
				<TR>
					<td class=texto1>Código negociador: <% = vContIdentArq %></td>
				</tr>
			</table> 
			<br>
			<table width="620" cellpadding="0" cellspacing="0" align="center" border=0 ID="Table27">
				<TR>
					<td class=texto1>Débito referente ao(s) contrato(s)/parcela(s):</td>
				</tr>
				<%
				vContr = ""
				Set rsTemp2 = BD.Execute("SELECT TRAN_NumTitulo, TDOC_Descricao FROM Transacoes_de_Acordo t WITH (NOLOCK) LEFT JOIN Tipos_de_Documento td WITH (NOLOCK) ON t.TDOC_ID = td.TDOC_ID WHERE ACOR_ID = " & rsParcelas("ACOR_ID"))
				if rsTemp2.EOF then
					Set rsTemp2 = BD.Execute("SELECT TRAN_NumTitulo, TDOC_Descricao FROM Transacoes t WITH (NOLOCK) LEFT JOIN Tipos_de_Documento td WITH (NOLOCK) ON t.TDOC_ID = td.TDOC_ID WHERE CTRA_ID = " & rsParcelas("CTRA_ID"))
				end if
				Do While Not rsTemp2.EOF
					vContr = vContr & rsTemp2("TRAN_NumTitulo") & " " & rsTemp2("TDOC_Descricao") & ", "
					rsTemp2.MoveNext
				Loop
				vContr = Mid(vContr, 1, len(vContr) - 2)
				%>
				<TR>
					<td class=texto1><% = vContr%></td>
				</tr>
				<%if cdate(rsParcelas("PARC_Vencimento")) < date() then%>
				<tr>
					<td class=texto1>
		Informamos que até o momento não acusamos o pagamento de sua parcela vencida em  <% = rsParcelas("PARC_Vencimento")%>.
<br>
		Solicitamos a regularização com urgência para evitarmos o cancelamento do acordo e a inclusão nos orgãos de proteção ao crédito.
<br>
		NÃO PERCA AS FACILIDADES E DESCONTOS CONCEDIDOS , EFETUE O PAGAMENTO DE SUA PARCELA EM ATRASO O MAIS BREVE POSSÍVEL. 
<br>
	    * Caso o pagamento tenha sido efetuado , por favor desconsiderar aviso .
			</td>
				</tr>
				<%end if%>
				
			</table> 
		</td>
	</tr> 
	</table> 
	</TABLE>
    <table width="640" cellpadding="3" cellspacing="0" align="center" border=0 bordercolor=black ID="Table3">
     <tr>  
      	<td class=texto1><b>Em caso de dúvida ligue para <% = rsFilial2("FILI_Tel1") %> ou compareça no seguinte endereço:</b><br><b><% = rsFilial2("FILI_Endereco") & " - " & rsFilial2("FILI_Bairro") & " - " & rsFilial2("FILI_Cidade") & " - " & rsFilial2("FILI_Estado") & " - CEP: " & rsFilial2("FILI_CEP") %></b><br><b>De 2ª a 6ª feira: de 9:00 hs às 18:00 hs</b></td>
     </tr>
    </table>
    <br>
	<TABLE WIDTH="640" BORDER=1 CELLSPACING=0 CELLPADDING=1 ID="Table24">
		<TR>
   			<TD><FONT class=ct>Código do Documento</FONT><BR><FONT class=cn>&nbsp;<%=var_numerodoc%></FONT></TD>
			<TD valign=top><FONT class=ct>Espécie</FONT><BR><FONT class=cn>&nbsp;<%=var_textoespecie%></FONT></TD>
			<TD valign=top><FONT class=ct>Quantidade</FONT><BR><FONT class=cn>&nbsp;</FONT></TD>
			<TD bgcolor="#CCCCCC">
				<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0 ID="Table22">
					<TR><TD align=left><FONT class=ct>Valor do Documento </FONT></TD></TR>
						<TR><TD align=center><FONT class=cp><% = FormatNumber(var_valordocumento - vValorTaxaBoleto)%></FONT></TD></TR>
				</TABLE>
			</TD>
			<TD valign=top><FONT class=ct>Espécie Doc.</FONT><BR><FONT class=cn>&nbsp;</FONT></TD>
			<TD><FONT class=ct>C&oacute;digo Cedente</FONT><BR>
					<FONT align=center class=cn>&nbsp;<% = var_textocodigocedente %></FONT></TD>
		</TR>
	</TABLE>
	<TABLE WIDTH="640" CELLSPACING=0 CELLPADDING=0 BORDER=0 ID="Table23">
		<TR>
				<TD align=right><FONT class=ct>Autentica&ccedil;&atilde;o Mec&acirc;nica</FONT><BR></TD>
	</TR>
	</TABLE>
	<br>
	<%
	elseif Session("CodigoCliente") = 25 and (Request.QueryString("automatico") <> "" or Request.QueryString("atraso") <> "" or vTipoBoleto = 1) then
			if vCont_id = 100 or vCont_id = 104 then 
				DescBoletoEmpresarioZogbi()
			else
				DescBoletoEmpresario vCont_id
			end if	
	else
	%>
	<img src="../images/linha2.gif" border=0 width="640" height=1>
	<TABLE WIDTH="640" CELLSPACING=0 CELLPADDING=0 BORDER=0 ID="Table10">
  		<TR>
			<TD class=cp VALIGN=BOTTOM WIDTH=225><IMG SRC="../images/logobanco_<% = rsBanco("BANC_ID_Boleto")%>.gif"></TD>
			<TD ALIGN=RIGHT VALIGN=BOTTOM><FONT class=ld><B>RECIBO DO SACADO</B></FONT></TD>
		</TR>
	</TABLE>
	<TABLE WIDTH="640" BORDER=1 CELLSPACING=0 CELLPADDING=1 ID="Table11">
	<TR>
				<TD COLSPAN=2><FONT class=ct>Cedente</FONT><BR><FONT class=cp>&nbsp;<%=cons_cedente%></FONT></TD>
				<TD width=15%><FONT class=ct>Ag&ecirc;ncia / C&oacute;digo Cedente</FONT><BR>
						<FONT align=center class=cn>&nbsp;<% = var_textocodigocedente %></FONT></TD>
   				<TD width=15%><FONT class=ct>Nosso Número</FONT><BR><FONT class=cn>&nbsp;<% = var_textonossonumero %></FONT></TD>
				<TD width=20% bgcolor="#CCCCCC">
				<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0 ID="Table12">
						<TR><TD align=left><FONT class=ct>Vencimento</FONT></TD></TR>
							<TR><TD align=center><FONT class=cp><%=var_datavencimento%></FONT></TD></TR>
					</TABLE>
				</TD>
		</TR>
		<TR>
			<%
			'Alteração realizada para imressão de boletos Ponto Frio
			if vCont_id <> 58 then %>
   			
   			<TD COLSPAN=2><FONT class=ct>Sacado</FONT><BR><FONT class=cp>&nbsp;<%=var_sacado%></FONT></TD>
			<TD width=15%><FONT class=ct> Parcela / Plano</FONT><BR><FONT align=center class=cn>&nbsp;<%= var_parcelaplano%></FONT></TD>
			
			<%else%>

   			<TD COLSPAN=3><FONT class=ct>Sacado</FONT><BR><FONT class=cp>&nbsp;<%=var_sacado%></FONT></TD>
s
			
			<%end if%>
   			<TD width=15%><FONT class=ct>Número Documento</FONT><BR><FONT class=cn>&nbsp;<%=var_numerodoc%></FONT></TD>
				<TD width=20% bgcolor="#CCCCCC">
				<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0 ID="Table13">
						<TR><TD align=left><FONT class=ct>Valor do Documento </FONT></TD></TR>
							<TR><TD align=center><FONT class=cp><% if rsBanco("BANC_ID_Boleto") = 237 or vCONT_ID = 77 or vCONT_ID = 97 then Response.Write FormatNumber(var_valordocumento) else Response.Write FormatNumber(var_valordocumento - vValorTaxaBoleto)%></FONT></TD></TR>
					</TABLE>
				</TD>
		</TR>
		<TR>
   			<TD><FONT class=ct>Contrato</FONT><BR><FONT class=cp>&nbsp;<%if Session("CodigoCliente") = 23 and vCont_id = 53 then response.Write FormataContrato(rsParcelas("CTRA_Numero"))  & " - " & vDescCert   else response.Write var_contrato  end if%></FONT></TD>
   			<TD width=15%><FONT class=ct>(-) Desconto / Abatimento</FONT><BR><FONT class=cp><br></FONT></TD>
   			<TD width=15%><FONT class=ct>(+) Mora / Multa</FONT><BR><FONT class=cn><br></FONT></TD>
   			<TD width=15%><FONT class=ct>(+) Outros Acréscimos</FONT><BR><FONT class=cn>&nbsp;<% if rsBanco("BANC_ID_Boleto") = 237 or vCONT_ID = 77 or vCONT_ID = 97 or (rsBanco("BANC_ID_Boleto") = 409 and Session("CodigoCliente") = 20)  then Response.Write "&nbsp;" else Response.Write FormatNumber(vValorTaxaBoleto)%></FONT></TD>
   			<TD width=20% bgcolor="#CCCCCC"><FONT class=ct>(=) Valor Cobrado</FONT><BR><FONT class=cp><center><%if vJurosAtraso or vCONT_ID = 77 or rsBanco("BANC_ID_Boleto") = 237 or (rsBanco("BANC_ID_Boleto") = 409 and Session("CodigoCliente") = 20) then response.Write "&nbsp;" else Response.Write FormatNumber(var_valordocumento)%></center></FONT></TD>
		</TR>
	</TABLE>
	<TABLE WIDTH="640" CELLSPACING=0 CELLPADDING=0 BORDER=0 ID="Table14">
		<TR>
				<TD align=right><FONT class=ct>Autentica&ccedil;&atilde;o Mec&acirc;nica</FONT><BR></TD>
	</TR>
	</TABLE>
	<%
	end if
	%>
	<img src="../images/corte.gif" border=0 width="640"><br>
	<TABLE WIDTH="640" BORDER=0 CELLSPACING=0 CELLPADDING=0 ID="Table15">
	<tr>
			<td class=cp width=150><div align="left"><img src="../images/logobanco_<% = rsBanco("BANC_ID_Boleto")%>.gif"></div></td>
  			<td width=3 valign="bottom"><img height=22 src="../images/barra.gif" width=2 border=0></td>
	  		<td class=cpt  width=58 valign="bottom"><div align="center"><font class="bc"><%=cons_banco%>-<%=cons_dvbanco%></font></div></td>
  			<td width=3 valign="bottom"><img height=22 src="../images/barra.gif" width=2 border=0></td>
	  		<td class=ld align=right width=453 valign="bottom"><span class='ld'><p align="right">&nbsp;<%=var_linhadigitavel%></span></td>
	</tr>
	</TABLE>
	<TABLE WIDTH="640" BORDER=1 CELLSPACING=0 CELLPADDING=1 ID="Table16">
	<TR>
				<TD COLSPAN=5 WIDTH=500>
						<FONT class=ct>Local de Pagamento</FONT><BR>
						<FONT class=cp>&nbsp;<% = var_localpagamento%></FONT>
				</TD>
				<%
				if cons_banco = "399" then
				%>
				<TD align=left width=50>
				<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0 ID="Table18">
						<TR><TD align=left><FONT class=ct>Parcela</FONT></TD></TR>
						<TR><TD align=center><FONT class=cn><%=var_textoparcela%></FONT></TD></TR>
					</TABLE>
				</TD>
				<%
				end if
				%>
				<TD width=<% = var_tamcoluna%> bgcolor="#CCCCCC">
				<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0 ID="Table28">
						<TR><TD align=left><FONT class=ct>Vencimento</FONT></TD></TR>
						<TR><TD align=center><FONT class=cp><%=var_datavencimento%></FONT></TD></TR>
					</TABLE>
				</TD>
		</TR>
		<TR>
				<TD COLSPAN=5 WIDTH=500><FONT class=ct>Cedente</FONT><BR><FONT class=cn>&nbsp;<%=cons_cedente%></FONT></TD>
				<TD width=170 colspan=<% = var_colunas%>>
				<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0 ID="Table29">
						<TR><TD align=left><FONT class=ct>Ag&ecirc;ncia / C&oacute;digo Cedente</FONT></TD></TR>
							<TR><TD align=center><FONT class=cn><% = var_textocodigocedente %></FONT></TD></TR>
					</TABLE>
				</TD>
		</TR>
		<TR>
				<TD valign=top><FONT class=ct>Data Documento</FONT><BR><FONT class=cn>&nbsp;<%=var_datadocumento%></FONT></TD>
				<TD valign=top><FONT class=ct>Número Documento</FONT><BR><FONT class=cn>&nbsp;<%=var_numerodoc%></FONT></TD>
				<TD valign=top><FONT class=ct>Tipo Docu.</FONT><BR><FONT class=cn>&nbsp;<% if rsBanco("BANC_ID_Boleto") = 623 then Response.Write "DS" else Response.Write "RECIBO"%></FONT></TD>
				<TD valign=top><FONT class=ct>Aceite</FONT><BR><FONT class=cn>&nbsp;<% = var_textoaceite%></FONT></TD>
				<TD valign=top><FONT class=ct>Data Processamento</FONT><BR><FONT class=cn>&nbsp;<%=var_data%></FONT></TD>
				<TD width=170 colspan=<% = var_colunas%>>
				<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0 ID="Table30">
						<TR><TD align=left><FONT class=ct>Nosso Número</FONT></TD></TR>
							<TR><TD align=center><FONT class=cn><% = var_textonossonumero %></FONT></TD></TR>
</TABLE>
				</TD>
		</TR>
		<TR>
				<TD valign=top><FONT class=ct>Uso Banco</FONT><BR><FONT class=cn>&nbsp;<% = vUso_Banco%></FONT></TD>
				<TD valign=top><FONT class=ct>Carteira</FONT><BR><FONT class=cn>&nbsp;<%=cons_carteira%></FONT></TD>
				<TD valign=top><FONT class=ct>Espécie</FONT><BR><FONT class=cn>&nbsp;<%=var_textoespecie%></FONT></TD>
				<TD valign=top><FONT class=ct>Quantidade</FONT><BR><FONT class=cn>&nbsp;</FONT></TD>
				<TD valign=top><FONT class=ct>Valor</FONT><BR><FONT class=cn>&nbsp;</FONT></TD>
				<TD width=170 bgcolor="#CCCCCC" colspan=<% = var_colunas%>>
				<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0 ID="Table37">
						<TR><TD align=left><FONT class=ct>Valor do Documento </FONT></TD></TR>
							<TR><TD align=center><FONT class=cp>
							<%
								if rsBanco("BANC_ID_Boleto") = 237 or vCONT_ID = 77 or vCONT_ID = 97  then 
									Response.Write FormatNumber(var_valordocumento) 
								else 
									Response.Write FormatNumber(var_valordocumento - vValorTaxaBoleto)
								end if	
									%></FONT></TD></TR>
					</TABLE>
				</TD>
		</TR>
		<TR>
				<TH COLSPAN=5 ROWSPAN=4 valign=top align=LEFT ><FONT class=ct>Instru&ccedil;&otilde;es (Todas as informações deste bloqueto são de inteira responsabilidade do cedente)</FONT><BR>
 					<TABLE WIDTH="475" ALIGN=RIGHT CELLSPACING=0 CELLPADDING=0 BORDER=0 ID="Table38">
						<TR>
							<TD valign=top align=left>
								<FONT class=cn>
								<%=var_instrucoes%>
								<%if (vCONT_ID = 69 or vCONT_ID = 43) and (vTipoBoleto = 1 or vTipoBoleto = 3) then %>
									<br>
									<b>Sr. Cliente Caso tenha ocorrido mudança de endereço, favor comparecer à agência para atualização dos dados cadastrais para que o carnê com os próximos pagamentos chegue no endereço correto.</b>
								<%end if%>
								</FONT>
							</TD>
						</TR>
					</TABLE>
				</TH>
				<TD WIDTH=170 colspan=<% = var_colunas%>><FONT class=ct>(-) Desconto / Abatimento</FONT><BR><FONT class=cn3>&nbsp;</FONT></TD>
		</TR>
		<TR>
				<TD WIDTH=170 colspan=<% = var_colunas%>><FONT class=ct>(+) Mora / Multa</FONT><BR><FONT class=cn3>&nbsp;</FONT></TD>
		</TR>
		<TR>
				<TD WIDTH=170 colspan=<% = var_colunas%>><FONT class=ct>(+) Outros Acréscimos</FONT><BR><FONT class=cn><center>&nbsp;
				<% 
					if rsBanco("BANC_ID_Boleto") = 237 or (rsBanco("BANC_ID_Boleto") = 409 and Session("CodigoCliente") = 20) or vCONT_ID = 77 or vCONT_ID = 97 then 
						Response.Write "&nbsp;" 
					else 
						Response.Write FormatNumber(vValorTaxaBoleto)
					end if
					%></center></FONT></TD>
		</TR>
		<TR>
				<TD WIDTH=170 colspan=<% = var_colunas%> bgcolor="#CCCCCC">
					<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0 ID="Table39">
							<TR><TD align=left><FONT class=ct>(=) Valor Cobrado</FONT></TD></TR>
								<TR><TD align=center><FONT class=cp><%if vJurosAtraso or vCONT_ID = 77 or rsBanco("BANC_ID_Boleto") = 237 or (rsBanco("BANC_ID_Boleto") = 409 and Session("CodigoCliente") = 20) then response.Write "&nbsp;" else Response.Write FormatNumber(var_valordocumento)%></FONT></TD></TR>
					</TABLE>
				</TD>
		</TR>
		<TR>
				<%
				if Session("CodigoCliente") = 12 then
				%>
				<TD COLSPAN=<% = 5 + var_colunas%> valign=top>
							<TABLE WIDTH="638" ALIGN=RIGHT CELLSPACING=0 CELLPADDING=0 BORDER=0 ID="Table40">
								<TR>
									<TD valign=top align=left width=100><FONT class=ct>Sacado</FONT></td>
									<TD valign=top align=left>
										<FONT class=cn5>
										<%if (vCONT_ID = 69 or vCONT_ID = 43) and (vTipoBoleto = 1 or vTipoBoleto = 3) then %>

											<%=var_sacado%><BR>
											<%=var_endereco%>,&nbsp;<%=var_bairro%><BR>
											<%=var_cidade%>&nbsp;-&nbsp;<%=var_estado%>&nbsp;-&nbsp;<BR>
											CEP:&nbsp;<%=var_cep%>

										<%else%>
											
												<%=var_sacado%><BR>
												<%=var_endereco%>,&nbsp;<%=var_bairro%><BR>
												<%=var_cidade%>&nbsp;-&nbsp;<%=var_estado%>&nbsp;-&nbsp;CEP&nbsp;<%=var_cep%><BR>

										<%end if%>
										</FONT>
	 								</TD>
								</TR>
							</TABLE>
				<%
				'campo sacado para boletos impressos na Hargos
				elseif Session("CodigoCliente") = 20 then
				%>

				<TD COLSPAN=<% = 5 + var_colunas%> valign=top>
							<TABLE WIDTH="638" ALIGN=RIGHT CELLSPACING=0 CELLPADDING=0 BORDER=0 ID="Table44">
								<TR>
									<TD valign=top align=left width=100><FONT class=ct>Sacado </FONT></td>
									<TD valign=top align=left>
										<FONT class=cn5>
										<%if (vCONT_ID = 69 or vCONT_ID = 43) and (vTipoBoleto = 1 or vTipoBoleto = 3) then %>

											<%=var_sacado%><BR>
											<%=var_endereco%>,&nbsp;<%=var_bairro%><BR>
											<%=var_cidade%>&nbsp;-&nbsp;<%=var_estado%>&nbsp;-&nbsp;<BR>
											CEP:&nbsp;<%=var_cep%>

										<%else%>
											
												<%=var_sacado%><BR>
												<%=var_endereco%>,&nbsp;<%=var_bairro%><BR>
												CEP:&nbsp;<%=var_cep%>&nbsp;-&nbsp;<%=var_cidade%>&nbsp;-&nbsp;<%=var_estado%><BR>

										<%end if%>
										</FONT>
	 								</TD>
								</TR>
							</TABLE>
				
				<%
				else
				%>
				<TD COLSPAN=<% = 5 + var_colunas%> valign=top>
						<FONT class=ct>Sacado</FONT><BR>
							<TABLE WIDTH="560" ALIGN=RIGHT CELLSPACING=0 CELLPADDING=0 BORDER=0 ID="Table41">
								<TR>
									<TD valign=top align=left>
										<FONT class=cn>
										<%if (vCONT_ID = 69 or vCONT_ID = 43) and (vTipoBoleto = 1 or vTipoBoleto = 3) then %>
										
											<%=var_sacado%><BR>
											<%=var_endereco%>,&nbsp;<%=var_bairro%>&nbsp;-&nbsp;<%=var_cidade%>&nbsp;-&nbsp;<%=var_estado%><BR>
											CEP:&nbsp;<%=var_cep%>
										<%else%>
										
											<%=var_sacado%><BR>
											<%=var_endereco%>,&nbsp;<%=var_bairro%>&nbsp;-&nbsp;<%=var_cep%>&nbsp;-&nbsp;<%=var_cidade%>&nbsp;-&nbsp;<%=var_estado%><BR>
										<%end if%>
											<!--<%=var_cpfcnpj%><BR>-->
										</FONT>
	 								</TD>
								</TR>
							</TABLE>
				<%
				end if
				%>
				</TD>
		</TR>
	</TABLE>
	<TABLE WIDTH="640" CELLSPACING=0 CELLPADDING=0 BORDER=0 ID="Table42">
		<TR>
				<TD class=ct align=right>
  					<div align="right">Autenticação Mecânica - <b class="cp">Ficha de Compensação</b></div>
				</TD>
		</TR>
		<TR>
				<TD align=left>
						<%
							call wbarcode(var_codigobarras)
							'response.Write GeraBarraTexto("23791272000000102404130060007222400500201090")
						%>
				</TD>
		</TR>
	</TABLE>
<%
	rsParcelas.MoveNext
	if vQtdBoletos = 2 and Not rsParcelas.EOF then
		vQtdboletos = 0
		%>
		<br class="pb">
		<%		
	elseif (vCONT_ID = 74 or vCONT_ID = 75) and (Request.QueryString("automatico") <> "" or Request.QueryString("atraso") <> "" or vTipoBoleto = 1) and Not rsParcelas.EOF then
		vQtdboletos = 0
		%>
		<br class="pb">
		<%
	elseif Session("CodigoCliente") = 25 and (Request.QueryString("automatico") <> "" or Request.QueryString("atraso") <> "" or vTipoBoleto = 1) and Not rsParcelas.EOF then
		vQtdboletos = 0
		%>
		<br class="pb">
		<%		
	elseif Not rsParcelas.EOF then
		%>
		<br>
		<br>
		<%
	end if
Loop

if Request.QueryString("atraso") <> "" then
	BD.Execute("DELETE FROM Boletos_Atraso_Temp WHERE FUNC_ID = " & Session("FUNC_ID"))
end if
%>
</CENTER>
</body>
</HTML>
<script language=javascript>
	document.getElementById("aguarde").style.display = "none";
</script>
<%
Dim  HoraIniRel9586325 : HoraIniRel9586325 = now

Function GeraBarraTexto(CodBarras)
	Dim Tipo(102)
	Dim vString, a, vTr, vConv, x
	Tipo(1) = "Jj"
	Tipo(2) = "Dj"
	Tipo(3) = "Li"
	Tipo(4) = "Bx"
	Tipo(5) = "Jw"
	Tipo(6) = "Dw"
	Tipo(7) = "Bn"
	Tipo(8) = "Jm"
	Tipo(9) = "Dm"
	Tipo(10) = "Qv"
	Tipo(11) = "Wd"
	Tipo(12) = "Sd"
	Tipo(13) = "Yc"
	Tipo(14) = "Qt"
	Tipo(15) = "Ws"
	Tipo(16) = "Ss"
	Tipo(17) = "Qh"
	Tipo(18) = "Wg"
	Tipo(19) = "Sg"
	Tipo(20) = "Ev"
	Tipo(21) = "Md"
	Tipo(22) = "Gd"
	Tipo(23) = "Oc"
	Tipo(24) = "Et"
	Tipo(25) = "Ms"
	Tipo(26) = "Gs"
	Tipo(27) = "Eh"
	Tipo(28) = "Mg"
	Tipo(29) = "Gg"
	Tipo(30) = "Uu"
	Tipo(31) = "[b"
	Tipo(32) = "Vb"
	Tipo(33) = "\a"
	Tipo(34) = "Ur"
	Tipo(35) = "[q"
	Tipo(36) = "Vq"
	Tipo(37) = "Uf"
	Tipo(38) = "[e"
	Tipo(39) = "Ve"
	Tipo(40) = "Bv"
	Tipo(41) = "Jd"
	Tipo(42) = "Dd"
	Tipo(43) = "Lc"
	Tipo(44) = "Bt"
	Tipo(45) = "Js"
	Tipo(46) = "Ds"
	Tipo(47) = "Bh"
	Tipo(48) = "Jg"
	Tipo(49) = "Dg"
	Tipo(50) = "Ru"
	Tipo(51) = "Xb"
	Tipo(52) = "Tb"
	Tipo(53) = "Za"
	Tipo(54) = "Rr"
	Tipo(55) = "Xq"
	Tipo(56) = "Tq"
	Tipo(57) = "Rf"
	Tipo(58) = "Xe"
	Tipo(59) = "Te"
	Tipo(60) = "Fu"
	Tipo(61) = "Nb"
	Tipo(62) = "Hb"
	Tipo(63) = "Pa"
	Tipo(64) = "Fr"
	Tipo(65) = "Nq"
	Tipo(66) = "Hq"
	Tipo(67) = "Ff"
	Tipo(68) = "Ne"
	Tipo(69) = "He"
	Tipo(70) = "A|"
	Tipo(71) = "Il"
	Tipo(72) = "Cl"
	Tipo(73) = "Kk"
	Tipo(74) = "Az"
	Tipo(75) = "Iy"
	Tipo(76) = "Cy"
	Tipo(77) = "Ap"
	Tipo(78) = "Io"
	Tipo(79) = "Co"
	Tipo(80) = "Q{"
	Tipo(81) = "Wj"
	Tipo(82) = "Sj"
	Tipo(83) = "Yi"
	Tipo(84) = "Qx"
	Tipo(85) = "Ww"
	Tipo(86) = "Sw"
	Tipo(87) = "Qn"
	Tipo(88) = "Wm"
	Tipo(89) = "Sm"
	Tipo(90) = "E{"
	Tipo(91) = "Mj"
	Tipo(92) = "Gj"
	Tipo(93) = "Oi"
	Tipo(94) = "Ex"
	Tipo(95) = "Mw"
	Tipo(96) = "Gw"
	Tipo(97) = "En"
	Tipo(98) = "Mm"
	Tipo(99) = "Gm"
	Tipo(100) = "B{"
	Tipo(101) = "("
	Tipo(102) = ")"
	
	vString = Tipo(101)
	vTr = len(CodBarras)
	if vTr mod 2 <> 0 then
		CodBarras = "0" & Trim(CodBarras)
	end if
	vConv = Trim(CodBarras)
	For x = 1 to len(vConv) step 2
		a = CInt(Mid(vConv, x, 2))
		if a < 1 then
			a = 100
		end if
		vString = Trim(vString) & Trim(Tipo(a))
	Next
	vString = Trim(vString) & Trim(Tipo(102))
	GeraBarraTexto = Trim(vString)
End Function

Function ValorIOF(SomatorioOrigens, QtdParc)
	if QtdParc = 1 then
		ValorIOF = SomatorioOrigens * 0.001241488
	elseif QtdParc = 2 then
		ValorIOF = SomatorioOrigens * 0.002482975
	elseif QtdParc = 3 then
		ValorIOF = SomatorioOrigens * 0.003724463
	elseif QtdParc = 4 then
		ValorIOF = SomatorioOrigens * 0.004965951
	elseif QtdParc = 5 then
		ValorIOF = SomatorioOrigens * 0.006207439
	elseif QtdParc = 6 then
		ValorIOF = SomatorioOrigens * 0.007448926
	elseif QtdParc = 7 then
		ValorIOF = SomatorioOrigens * 0.008690414
	elseif QtdParc = 8 then
		ValorIOF = SomatorioOrigens * 0.009931902
	elseif QtdParc = 9 then
		ValorIOF = SomatorioOrigens * 0.011173389
	elseif QtdParc = 10 then
		ValorIOF = SomatorioOrigens * 0.012414877
	elseif QtdParc = 11 then
		ValorIOF = SomatorioOrigens * 0.013656365
	elseif QtdParc = 12 then
		ValorIOF = SomatorioOrigens * 0.014897853
	elseif QtdParc = 13 then
		ValorIOF = SomatorioOrigens * 0.016139340
	elseif QtdParc = 14 then
		ValorIOF = SomatorioOrigens * 0.017380828
	elseif QtdParc = 15 then
		ValorIOF = SomatorioOrigens * 0.018622316
	elseif QtdParc = 16 then
		ValorIOF = SomatorioOrigens * 0.019863803
	elseif QtdParc = 17 then
		ValorIOF = SomatorioOrigens * 0.021105291
	elseif QtdParc = 18 then
		ValorIOF = SomatorioOrigens * 0.022346779
	elseif QtdParc = 19 then
		ValorIOF = SomatorioOrigens * 0.023588267
	elseif QtdParc = 20 then
		ValorIOF = SomatorioOrigens * 0.024829754
	elseif QtdParc = 21 then
		ValorIOF = SomatorioOrigens * 0.026071242
	elseif QtdParc = 22 then
		ValorIOF = SomatorioOrigens * 0.027312730
	elseif QtdParc = 23 then
		ValorIOF = SomatorioOrigens * 0.028554217
	elseif QtdParc = 24 then
		ValorIOF = SomatorioOrigens * 0.029795705
	elseif QtdParc = 25 then
		ValorIOF = SomatorioOrigens * 0.031037193
	elseif QtdParc = 26 then
		ValorIOF = SomatorioOrigens * 0.032278681
	elseif QtdParc = 27 then
		ValorIOF = SomatorioOrigens * 0.033520168
	elseif QtdParc = 28 then
		ValorIOF = SomatorioOrigens * 0.034761656
	elseif QtdParc = 29 then
		ValorIOF = SomatorioOrigens * 0.036003144
	elseif QtdParc = 30 then
		ValorIOF = SomatorioOrigens * 0.037244631
	elseif QtdParc = 31 then
		ValorIOF = SomatorioOrigens * 0.038486119
	elseif QtdParc = 32 then
		ValorIOF = SomatorioOrigens * 0.039727607
	elseif QtdParc = 33 then
		ValorIOF = SomatorioOrigens * 0.040969095
	elseif QtdParc = 34 then
		ValorIOF = SomatorioOrigens * 0.042210582
	elseif QtdParc = 35 then
		ValorIOF = SomatorioOrigens * 0.043452070
	elseif QtdParc = 36 then
		ValorIOF = SomatorioOrigens * 0.044693558
	elseif QtdParc = 37 then
		ValorIOF = SomatorioOrigens * 0.045935046
	elseif QtdParc = 38 then
		ValorIOF = SomatorioOrigens * 0.047176533
	elseif QtdParc = 39 then
		ValorIOF = SomatorioOrigens * 0.048418021
	elseif QtdParc = 40 then
		ValorIOF = SomatorioOrigens * 0.049659509
	elseif QtdParc = 41 then
		ValorIOF = SomatorioOrigens * 0.050900996
	elseif QtdParc = 42 then
		ValorIOF = SomatorioOrigens * 0.052142484
	elseif QtdParc = 43 then
		ValorIOF = SomatorioOrigens * 0.053383972
	elseif QtdParc = 44 then
		ValorIOF = SomatorioOrigens * 0.054625460
	elseif QtdParc = 45 then
		ValorIOF = SomatorioOrigens * 0.055866947
	elseif QtdParc = 46 then
		ValorIOF = SomatorioOrigens * 0.057108435
	elseif QtdParc = 47 then
		ValorIOF = SomatorioOrigens * 0.058349923
	elseif QtdParc = 48 then
		ValorIOF = SomatorioOrigens * 0.059591410
	else
		ValorIOF = 0
	end if
End Function

Function GeraSQLScript(Tabela, Condicao, CampoTroca0, ValorTroca0, CampoTroca1, ValorTroca1, CampoTroca2, ValorTroca2, CampoTroca3, ValorTroca3, CampoTroca4, ValorTroca4)
	Dim RSTemp, QTDCampo, Idx ,vValue, SQLScript
	SQLScript = ""
	Set RSTemp = BD.Execute("Select * FROM " & Tabela & " WITH (NOLOCK) WHERE " & Condicao )
	'If Not RSTemp.EOF Then
	'	SQLScript = "INSERT INTO " & Tabela & " ("
	'	For idx = 0 To RSTemp.Fields.Count - 1
	'		SQLScript = SQLScript & RSTemp.Fields(Idx).Name & ","
	'	Next		
	'	SQLScript = mid(SQLScript,1,len(SQLScript)-1) & ") Values ("  
	'	For idx = 0 To RSTemp.Fields.Count - 1
	'		vValue=""
	'		'Response.Write RSTemp.Fields(Idx).name & "(" & RSTemp.Fields(Idx).value & ")=<BR>"			
	'		If     Ucase(RSTemp.Fields(Idx).Name) = UCase(TRIM(CampoTroca0)) Then 
	'			vValue = ValorTroca0
	'		ElseIf Ucase(RSTemp.Fields(Idx).Name) = UCase(TRIM(CampoTroca1)) Then 
	'			vValue = ValorTroca1
	'		ElseIf Ucase(RSTemp.Fields(Idx).Name) = UCase(TRIM(CampoTroca2)) Then 
	'			vValue = ValorTroca2
	'		ElseIf Ucase(RSTemp.Fields(Idx).Name) = UCase(TRIM(CampoTroca3)) Then 
	'			vValue = ValorTroca3
	'		ElseIf Ucase(RSTemp.Fields(Idx).Name) = UCase(TRIM(CampoTroca4)) Then 
	'			vValue = ValorTroca4
	'		Elseif Ucase(RSTemp.Fields(Idx).Name) = UCase("PARC_NumDocumento") Then
	'			IF Not IsNull(RSTemp.Fields(Idx).value) then 
	'				vValue = RSTemp.Fields(Idx).value
	'			Else
	'				vValue = "Null"
	'			End if
	'		Elseif Ucase(RSTemp.Fields(Idx).Name) = UCase("BOAV_NumDocumento") Then
	'			IF Not IsNull(RSTemp.Fields(Idx).value) then 
	'				vValue = RSTemp.Fields(Idx).value
	'			Else
	'				vValue = "Null"
	'			End if
	'		Elseif Ucase(RSTemp.Fields(Idx).Name) = UCase("PAGA_NumDocumento") Then
	'			IF Not IsNull(RSTemp.Fields(Idx).value) then 
	'				vValue = RSTemp.Fields(Idx).value
	'			Else
	'				vValue = "Null"
	'			End if
	'		Elseif Ucase(RSTemp.Fields(Idx).Name) = UCase("CTRA_CodigoCliente") Then
	'			IF Not IsNull(RSTemp.Fields(Idx).value) then 
	'				vValue = RSTemp.Fields(Idx).value
	'			Else
	'				vValue = "Null"
	'			End if
	'		Else
	'			'RESPONSE.Write RSTemp.Fields(Idx).name
	'			Select Case  TypeName(RSTemp.Fields(Idx).value)
	'				Case "Null"
	'					vValue = "Null"
	'				Case "Boolean"
	'					IF RSTemp.Fields(Idx).value = True Then vValue = "1" Else vValue = "0"
	'				Case "Long" , "Integer"
	'					'vValue = "'" & Replace(RSTemp.Fields(idx).value,"'","''") & "'"	
	'					vValue = RSTemp.Fields(Idx).value
	'				Case "Currency" , "Double", "Single","Decimal"
	'					vValue = Replace(Replace(RSTemp.Fields(Idx).value, "." , ""), "," , "." )
	'				Case "Date"
	'					vValue = "'" & RSTemp.Fields(idx).value & "'"	
	'				Case "String"
	'					vValue = "'" & Replace(RSTemp.Fields(idx).value,"'","''") & "'"	
	'				Case Else
	'					Response.Write "***********" & RSTemp.Fields(Idx).name & "(" & RSTemp.Fields(Idx).value  & ")=" & TypeName(RSTemp.Fields(Idx).value) & "<br><br>"
	'			End Select
	'		End If			
	'		SQLScript = SQLScript & vValue  & ","
			'response.Write SQLScript
			'Response.Write "<BR>" & chr(13) & chr(10)
	'	Next
	'	SQLScript = Mid(SQLScript,1,Len(SQLScript)-1) & ")"  
	'End if
	RSTemp.Close
	Set RSTemp = Nothing
	'response.Write SQLScript & "<BR>" 
	GeraSQLScript = SQLScript
	'RESPONSE.Write "---------------------------------------<BR>"
End Function

Function GeraSQLScript2(Tabela, Condicao, CampoTroca0, ValorTroca0, CampoTroca1, ValorTroca1, CampoTroca2, ValorTroca2, CampoTroca3, ValorTroca3, CampoTroca4, ValorTroca4)
	Dim RSTemp, QTDCampo, Idx ,vValue, SQLScript
	SQLScript = ""
	Set RSTemp = BD.Execute("Select * FROM " & Tabela & " WITH (NOLOCK) WHERE " & Condicao )
'	If Not RSTemp.EOF Then
'		SQLScript = "INSERT INTO " & Tabela & " ("
'		For idx = 1 To RSTemp.Fields.Count - 1
'			SQLScript = SQLScript & RSTemp.Fields(Idx).Name & ","
'		Next		
'		SQLScript = mid(SQLScript,1,len(SQLScript)-1) & ") Values ("  
'		For idx = 1 To RSTemp.Fields.Count - 1
'			vValue=""
'			'Response.Write RSTemp.Fields(Idx).name & "(" & RSTemp.Fields(Idx).value & ")="			
'			If     Ucase(RSTemp.Fields(Idx).Name) = UCase(TRIM(CampoTroca0)) Then 
'				vValue = ValorTroca0
'			ElseIf Ucase(RSTemp.Fields(Idx).Name) = UCase(TRIM(CampoTroca1)) Then 
'				vValue = ValorTroca1
'			ElseIf Ucase(RSTemp.Fields(Idx).Name) = UCase(TRIM(CampoTroca2)) Then 
'				vValue = ValorTroca2
'			ElseIf Ucase(RSTemp.Fields(Idx).Name) = UCase(TRIM(CampoTroca3)) Then 
'				vValue = ValorTroca3
'			ElseIf Ucase(RSTemp.Fields(Idx).Name) = UCase(TRIM(CampoTroca4)) Then 
'				vValue = ValorTroca4
'			Elseif Ucase(RSTemp.Fields(Idx).Name) = UCase("PARC_NumDocumento") Then
'				IF Not IsNull(RSTemp.Fields(Idx).value) then 
'					vValue = RSTemp.Fields(Idx).value
'				Else
'					vValue = "Null"
'				End if
'			Elseif Ucase(RSTemp.Fields(Idx).Name) = UCase("BOAV_NumDocumento") Then
'				IF Not IsNull(RSTemp.Fields(Idx).value) then 
'					vValue = RSTemp.Fields(Idx).value
'				Else
'					vValue = "Null"
'				End if
'			Elseif Ucase(RSTemp.Fields(Idx).Name) = UCase("PAGA_NumDocumento") Then
'				IF Not IsNull(RSTemp.Fields(Idx).value) then 
'					vValue = RSTemp.Fields(Idx).value
'				Else
'					vValue = "Null"
'				End if
'			Else
'				'RESPONSE.Write RSTemp.Fields(Idx).name
'				Select Case  TypeName(RSTemp.Fields(Idx).value)
'					Case "Null"
'						vValue = "Null"
'					Case "Boolean"
'						IF RSTemp.Fields(Idx).value = True Then vValue = "1" Else vValue = "0"
'					Case "Long" , "Integer"
'						'vValue = "'" & Replace(RSTemp.Fields(idx).value,"'","''") & "'"	
'						vValue = RSTemp.Fields(Idx).value
'					Case "Currency" , "Double", "Single","Decimal"
'						vValue = Replace(Replace(RSTemp.Fields(Idx).value, "." , ""), "," , "." )
'					Case "Date"
'						vValue = "'" & RSTemp.Fields(idx).value & "'"	
'					Case "String"
'						vValue = "'" & Replace(RSTemp.Fields(idx).value,"'","''") & "'"	
'					Case Else
'						Response.Write "***********" & RSTemp.Fields(Idx).name & "(" & RSTemp.Fields(Idx).value  & ")=" & TypeName(RSTemp.Fields(Idx).value) & "<br><br>"
'				End Select
'			End If			
'			SQLScript = SQLScript & vValue  & ","
'			'Response.Write "<BR>" & chr(13) & chr(10)
'		Next
'		SQLScript = Mid(SQLScript,1,Len(SQLScript)-1) & ")"  
'	End if
	RSTemp.Close
	Set RSTemp = Nothing
'	GeraSQLScript2 = SQLScript
	'RESPONSE.Write "---------------------------------------<BR>"
End Function

Sub DuplicaContrato(Contrato, Devedor, TitulosAcordo, Contratante, Carteira, FilialDestino)
	'Duplica o contrato com um sequencial no final de 3 dígitos
	'O sequencial do contrato é gravado no campo CTRA_SequencialContrato
	'Quando duplicar o contrato, copiar todos os títulos que não fazem parte do acordo
	'Deletar as transações transferidas para o novo contrato
	'Colocar o contrato para o mesmo recuperador
	'Caso o contrato só tenha títulos a vencer, devolver o contrato
	DIM Data_OLD
	DIM CONT_ID
	DIM BORD_ID_NEW
	DIM DEVE_ID_OLD, DEVE_ID_NEW
	DIM CTRA_ID_OLD, CTRA_ID_NEW
	DIM CART_ID_OLD, CART_ID_NEW
	DIM CART_NO_OLD, CART_NO_NEW
	DIM FILI_ID_OLD, FILI_ID_NEW, FILI_ID_FUN
	DIM FILI_NO_OLD, FILI_NO_NEW
	DIM DAD0_ID_OLD, DAD0_ID_NEW
	DIM DAD1_ID_OLD, DAD1_ID_NEW
	DIM DAD2_ID_OLD, DAD2_ID_NEW
	DIM Motivo,      Prazo_Devolucao
	Dim Andamento
	Dim BD, SQLScript, RSTabela0, RSTabela1, RSTabela2
	DIM Condicao
	Dim rsContrato, rsContrato2, vSeqContrato, CTRA_NUMERO_OLD, CTRA_NUMERO_NEW, rsT, rsT2
	
	if TitulosAcordo <> "" then
		Set BD = Server.CreateObject("ADODB.Connection")
		BD.Open Application("Conexao")

		DEVE_ID_OLD = Devedor
		CTRA_ID_OLD = Contrato
		CONT_ID = Contratante
		
		Set rsT = BD.Execute("SELECT TRAN_ID FROM Transacoes WITH (NOLOCK) WHERE CTRA_ID = " & Contrato & " AND TRAN_ID NOT IN (" & TitulosAcordo & ")")
		if not rsT.EOF then
			Set rsContrato = BD.Execute("SELECT CTRA_SequencialContrato, CTRA_Numero FROM Contratos WITH (NOLOCK) WHERE CTRA_ID = " & CTRA_ID_OLD)
			if isnull(rsContrato("CTRA_SequencialContrato")) then
				CTRA_NUMERO_OLD = rsContrato("CTRA_Numero")
			else
				CTRA_NUMERO_OLD = Left(rsContrato("CTRA_Numero"), len(rsContrato("CTRA_Numero")) - 3)
			end if
			Set rsContrato = BD.Execute("SELECT MAX(CTRA_SequencialContrato) CTRA_SequencialContrato FROM Contratos c WITH (NOLOCK) JOIN Carteiras ca WITH (NOLOCK) ON c.CART_ID = ca.CART_ID WHERE CONT_ID = " & Contratante & " AND CTRA_Numero like '" & CTRA_NUMERO_OLD & "%'")
			if isnull(rsContrato("CTRA_SequencialContrato")) then
				vSeqContrato = 2
			else
				vSeqContrato = rsContrato("CTRA_SequencialContrato") + 1
			end if
			CTRA_NUMERO_NEW = CTRA_NUMERO_OLD & Right("000" & vSeqContrato, 3) 

			'**********************************************************************************************************************
			' Devedores
			'**********************************************************************************************************************
			DEVE_ID_NEW = RetornaProximoNumero("Devedores", "DEVE_ID", Session("FILI_ID")) 
			SQLScript = GeraSQLScript("Devedores","DEVE_ID = " & DEVE_ID_OLD,"DEVE_ID", DEVE_ID_NEW,"","","","","","","","")
			BD.Execute SQLScript 
			'PreencheLOG_Contratante SQLScript , FilialDestino, CONT_ID
			
			'**********************************************************************************************************************
			' Contratos
			'**********************************************************************************************************************
			CTRA_ID_NEW = RetornaProximoNumero("Contratos", "CTRA_ID", Session("FILI_ID")) 
			SQLScript = GeraSQLScript("Contratos","CTRA_ID=" & CTRA_ID_OLD, "CTRA_ID",CTRA_ID_NEW,"CTRA_Numero","'" & CTRA_NUMERO_NEW & "'","DEVE_ID",DEVE_ID_NEW,"CTRA_SequencialContrato",vSeqContrato,"","")
			BD.Execute SQLScript  
			'PreencheLOG_Contratante SQLScript , FilialDestino, CONT_ID

			'SQLScript = "UPDATE Contratos SET CTRA_NumeroAcordoContratante = NULL WHERE CTRA_ID = " & CTRA_ID_NEW
			'BD.Execute SQLScript  
			'PreencheLOG_Contratante SQLScript , FilialDestino, CONT_ID

			'**********************************************************************************************************************
			' Andamento_da_Cobranca / Telefones do Devedor informados no Andamento da Cobrança
			'**********************************************************************************************************************
			DIM TDEV_ID_OLD, TDEV_ID_NEW
			DIM TDEV_ID_LST :	TDEV_ID_LST = "" 
			Set RSTabela0 = BD.Execute("SELECT * FROM Andamento_da_Cobranca WITH (NOLOCK) WHERE CTRA_ID = " & CTRA_ID_OLD )
			IF NOT RSTabela0.EOF then
				DO WHILE NOT RSTabela0.EOF
					'**********************************************************************************************************
					' Telefone do Devedor
					'**********************************************************************************************************
					IF Not IsNull(RSTabela0("TDEV_ID")) then
						TDEV_ID_OLD = RSTabela0("TDEV_ID")
						TDEV_ID_NEW = RetornaProximoNumero("Telefones_do_Devedor", "TDEV_ID", Session("FILI_ID"))
						TDEV_ID_LST = TDEV_ID_LST & TDEV_ID_OLD & ","
						SQLScript = GeraSQLScript("Telefones_do_Devedor","TDEV_ID =" & TDEV_ID_OLD, "TDEV_ID", TDEV_ID_NEW, "DEVE_ID", DEVE_ID_NEW,"","","","","","")
						if SQLScript <> "" then
							BD.Execute SQLScript
							'PreencheLOG_Contratante SQLScript , FilialDestino, CONT_ID
						end if
						'PreencheLOG_Contratante SQLScript , FilialDestino, CONT_ID
					Else
						TDEV_ID_NEW = "NULL"
					End if
					'**********************************************************************************************************
					if isnull(RSTabela0("FUNC_ID")) then
						'SQLScript = "Insert Into Andamento_da_Cobranca Values (" & CTRA_ID_NEW & ",'" & RSTabela0("ANCO_DataHora") & "'," & RSTabela0("STAC_ID") & "," & TDEV_ID_NEW & ",'" & Replace(RSTabela0("ANCO_Descricao"), "'", "''") & "',NULL)"
					else
						'SQLScript = "Insert Into Andamento_da_Cobranca Values (" & CTRA_ID_NEW & ",'" & RSTabela0("ANCO_DataHora") & "'," & RSTabela0("STAC_ID") & "," & TDEV_ID_NEW & ",'" & Replace(RSTabela0("ANCO_Descricao"), "'", "''") & "'," & RSTabela0("FUNC_ID") & ")"
					end if
					'Response.Write SQLScript & "<br>"
					'BD.Execute SQLScript
					'PreencheLOG_Contratante SQLScript , FilialDestino, CONT_ID
					RSTabela0.MoveNext
				LOOP
			END IF
			RSTabela0.Close

			'**********************************************************************************************************************
			' Telefones_do_Devedor
			'**********************************************************************************************************************
			IF TDEV_ID_LST <> "" then Condicao = " AND TDEV_ID not in (" & Mid(TDEV_ID_LST,1,len(TDEV_ID_LST) -1) & ")"
			Set RSTabela0 = BD.Execute("SELECT TDEV_ID FROM Telefones_do_Devedor WITH (NOLOCK) WHERE DEVE_ID = " & DEVE_ID_OLD & Condicao)
			IF NOT RSTabela0.EOF then
				DO WHILE NOT RSTabela0.EOF
					DAD0_ID_OLD = RSTabela0(0)
					DAD0_ID_NEW = RetornaProximoNumero("Telefones_do_Devedor", "TDEV_ID", Session("FILI_ID")) 
					SQLScript = GeraSQLScript("Telefones_do_Devedor","TDEV_ID =" & DAD0_ID_OLD, "TDEV_ID", DAD0_ID_NEW, "DEVE_ID", DEVE_ID_NEW,"","","","","","")
					BD.Execute SQLScript
					'PreencheLOG_Contratante SQLScript , FilialDestino, CONT_ID
					RSTabela0.MoveNext
				LOOP
			END IF
			RSTabela0.Close

			'**********************************************************************************************************************
			' Endereco_Devedor
			'**********************************************************************************************************************
			Set RSTabela0 = BD.Execute("SELECT EDEV_ID FROM Endereco_Devedor WITH (NOLOCK) WHERE DEVE_ID = " & DEVE_ID_OLD)
			IF NOT RSTabela0.EOF then
				DO WHILE NOT RSTabela0.EOF
					DAD0_ID_OLD = RSTabela0(0)
					DAD0_ID_NEW = RetornaProximoNumero("Endereco_Devedor", "EDEV_ID", Session("FILI_ID")) 
					SQLScript = GeraSQLScript("Endereco_Devedor","EDEV_ID =" & DAD0_ID_OLD, "EDEV_ID", DAD0_ID_NEW, "DEVE_ID", DEVE_ID_NEW,"","","","","","")
					BD.Execute SQLScript
					'PreencheLOG_Contratante SQLScript , FilialDestino, CONT_ID
					RSTabela0.MoveNext
				LOOP
			END IF
			RSTabela0.Close

			'**********************************************************************************************************************
			' Avalistas_Referencias / Telefones_do_Avalista
			'**********************************************************************************************************************
			Set RSTabela0 = BD.Execute("SELECT AVRE_ID FROM Avalistas_Referencias WITH (NOLOCK) WHERE DEVE_ID = " & DEVE_ID_OLD)
			IF NOT RSTabela0.EOF then
				DO WHILE NOT RSTabela0.EOF
					DAD0_ID_OLD = RSTabela0(0)
					DAD0_ID_NEW = RetornaProximoNumero("Avalistas_Referencias", "AVRE_ID", Session("FILI_ID")) 
					SQLScript = GeraSQLScript("Avalistas_Referencias","AVRE_ID =" & DAD0_ID_OLD, "AVRE_ID", DAD0_ID_NEW, "DEVE_ID", DEVE_ID_NEW,"","","","","","")
					BD.Execute SQLScript
					'PreencheLOG_Contratante SQLScript , FilialDestino, CONT_ID
					'**********************************************************************************************************
					' Telefones_do_Avalista
					'**********************************************************************************************************
					Set RSTabela1 = BD.Execute("SELECT TAVA_ID FROM Telefones_do_Avalista WITH (NOLOCK) WHERE AVRE_ID = " & DAD0_ID_OLD )
						IF NOT RSTabela1.EOF then
							DO WHILE NOT RSTabela1.EOF
								DAD1_ID_OLD = RSTabela1(0)
								DAD1_ID_NEW = RetornaProximoNumero("Telefones_do_Avalista", "TAVA_ID", Session("FILI_ID")) 
								SQLScript = GeraSQLScript("Telefones_do_Avalista","TAVA_ID =" & DAD1_ID_OLD, "TAVA_ID", DAD1_ID_NEW, "AVRE_ID", DAD0_ID_NEW,"","","","","","")
								BD.Execute SQLScript
								'PreencheLOG_Contratante SQLScript , FilialDestino, CONT_ID
								RSTabela1.MoveNext
							LOOP
						END IF
					RSTabela1.Close
					'**********************************************************************************************************
					RSTabela0.MoveNext
				LOOP
			END IF
			RSTabela0.Close

			'**********************************************************************************************************************
			' REL_Contrato_Bordero
			'**********************************************************************************************************************
			Set RSTabela0 = BD.Execute("SELECT RCB.*, B.TBOR_ID FROM REL_Contrato_Bordero RCB WITH (NOLOCK) JOIN Bordero B WITH (NOLOCK) ON RCB.BORD_ID = B.BORD_ID WHERE CTRA_ID = " & CTRA_ID_OLD)
			IF Not RSTabela0.eof Then
				Do While Not RSTabela0.Eof
					Select Case RSTabela0("TBOR_ID")
						Case "E" 'Obs.: Os tipos P e D não serão migrados 
							DAD0_ID_OLD = RSTabela0("BORD_ID")
							SQLScript = GeraSQLScript("REL_Contrato_Bordero","BORD_ID =" & DAD0_ID_OLD & " AND CTRA_ID = " & CTRA_ID_OLD, "CTRA_ID", CTRA_ID_NEW, "", "","","","","","","")
							BD.Execute SQLScript
							'PreencheLOG_Contratante SQLScript , FilialDestino, CONT_ID
					End Select
					RSTabela0.MoveNext
				Loop
			End IF

			'**********************************************************************************************************************
			' Dados_Adicionais_do_Contrato
			'**********************************************************************************************************************
			Set RSTabela0 = BD.Execute("SELECT * FROM Dados_Adicionais_do_Contrato WITH (NOLOCK) WHERE CTRA_ID = " & CTRA_ID_OLD)
			IF NOT RSTabela0.EOF then
				DO WHILE NOT RSTabela0.EOF
					SQLScript = GeraSQLScript("Dados_Adicionais_do_Contrato","CTRA_ID =" & CTRA_ID_OLD & " AND TDAD_ID = " & RSTabela0("TDAD_ID") & " AND DADO_Valor = '" & RSTabela0("DADO_Valor") &"'", "CTRA_ID", CTRA_ID_NEW,"","","","","","","","")
					BD.Execute SQLScript
					'PreencheLOG_Contratante SQLScript , FilialDestino, CONT_ID
					RSTabela0.MoveNext
				LOOP
			END IF
			RSTabela0.Close

			'**********************************************************************************************************************
			' Pagamentos_Avulsos / Transacoes vinculadas aos Pagamentos Avulsos migrados
			'**********************************************************************************************************************
			DIM TRAN_ID_PG_OLD, TRAN_ID_PG_NEW
			DIM TRAN_ID_GE_OLD, TRAN_ID_GE_NEW
			DIM TRAN_ID_LST :	TRAN_ID_LST = "" 
			Set RSTabela0 = BD.Execute("SELECT PAGA_ID, TRAN_ID_Paga, TRAN_ID_Gerada  FROM Pagamentos_Avulsos WITH (NOLOCK) WHERE CTRA_ID = " & CTRA_ID_OLD )
			IF NOT RSTabela0.EOF then
				DO WHILE NOT RSTabela0.EOF
					DAD0_ID_OLD = RSTabela0("PAGA_ID")
					DAD0_ID_NEW = RetornaProximoNumero("Pagamentos_Avulsos", "PAGA_ID", Session("FILI_ID")) 
					'**********************************************************************************************************
					' TRAN_ID_Paga
					'**********************************************************************************************************
					TRAN_ID_PG_OLD = RSTabela0("TRAN_ID_Paga")
					TRAN_ID_PG_NEW = RetornaProximoNumero("Transacoes", "TRAN_ID", Session("FILI_ID"))
					TRAN_ID_LST = TRAN_ID_LST & TRAN_ID_PG_OLD & ","
					SQLScript = GeraSQLScript("Transacoes","TRAN_ID =" & TRAN_ID_PG_OLD, "TRAN_ID", TRAN_ID_PG_NEW, "CTRA_ID", CTRA_ID_NEW,"","","","","","")
					BD.Execute SQLScript
					'PreencheLOG_Contratante SQLScript , FilialDestino, CONT_ID
					'**********************************************************************************************************
					' TRAN_ID_Gerada
					'**********************************************************************************************************
					TRAN_ID_GE_OLD = RSTabela0("TRAN_ID_Gerada")
					TRAN_ID_GE_NEW = RetornaProximoNumero("Transacoes", "TRAN_ID", Session("FILI_ID"))
					TRAN_ID_LST = TRAN_ID_LST & TRAN_ID_GE_OLD & ","
					SQLScript = GeraSQLScript("Transacoes","TRAN_ID =" & TRAN_ID_GE_OLD, "TRAN_ID", TRAN_ID_GE_NEW, "CTRA_ID", CTRA_ID_NEW,"","","","","","")
					BD.Execute SQLScript
					'PreencheLOG_Contratante SQLScript , FilialDestino, CONT_ID
					'**********************************************************************************************************
					SQLScript = GeraSQLScript("Pagamentos_Avulsos","PAGA_ID =" & DAD0_ID_OLD, "PAGA_ID", DAD0_ID_NEW, "CTRA_ID", CTRA_ID_NEW,"TRAN_ID_Paga",TRAN_ID_PG_NEW,"TRAN_ID_Gerada",TRAN_ID_GE_NEW,"","")
					BD.Execute SQLScript
					'PreencheLOG_Contratante SQLScript , FilialDestino, CONT_ID
					RSTabela0.MoveNext
				LOOP
			END IF
			RSTabela0.Close
			'Deleta os pagamentos avulsos
			'BD.Execute("DELETE FROM Pagamentos_Avulsos WHERE CTRA_ID = " & CTRA_ID_OLD )
			''PreencheLOG_Contratante "DELETE FROM Pagamentos_Avulsos WHERE CTRA_ID = " & CTRA_ID_OLD, FilialDestino, CONT_ID

			'**********************************************************************************************************************
			' Transacoes que não tiveram Pagamentos Avulsos
			'**********************************************************************************************************************
			Condicao = ""
			IF TRAN_ID_LST <> "" then Condicao = " AND TRAN_ID not in (" & MId(TRAN_ID_LST,1,len(TRAN_ID_LST) -1) & ")"
			Set RSTabela0 = BD.Execute("SELECT TRAN_ID FROM Transacoes WITH (NOLOCK) WHERE CTRA_ID = " & CTRA_ID_OLD & Condicao & " AND TRAN_ID NOT IN (" & TitulosAcordo & ")")
			IF NOT RSTabela0.EOF then
				DO WHILE NOT RSTabela0.EOF
					DAD0_ID_OLD = RSTabela0(0)
					DAD0_ID_NEW = RetornaProximoNumero("Transacoes", "TRAN_ID", Session("FILI_ID")) 
					SQLScript = GeraSQLScript("Transacoes","TRAN_ID =" & DAD0_ID_OLD, "TRAN_ID", DAD0_ID_NEW, "CTRA_ID", CTRA_ID_NEW,"","","","","","")
					BD.Execute SQLScript
					'PreencheLOG_Contratante SQLScript , FilialDestino, CONT_ID
					RSTabela0.MoveNext
				LOOP
			END IF
			RSTabela0.Close
			'Deleta as transações que não pertencem ao acordo
			'BD.Execute("DELETE FROM Transacoes WHERE CTRA_ID = " & CTRA_ID_OLD & " AND TRAN_ID NOT IN (" & TitulosAcordo & ")")
			''PreencheLOG_Contratante "DELETE FROM Transacoes WHERE CTRA_ID = " & CTRA_ID_OLD & " AND TRAN_ID NOT IN (" & TitulosAcordo & ")", FilialDestino, CONT_ID
			'Altera o vencimento do débito para a transação mais antiga
			'BD.Execute("UPDATE Contratos SET CTRA_VencDebito = (SELECT MIN(TRAN_Vencimento) FROM Transacoes WITH (NOLOCK) WHERE CTRA_ID = " & CTRA_ID_NEW & ") WHERE CTRA_ID = " & CTRA_ID_NEW)
			''PreencheLOG_Contratante "UPDATE Contratos SET CTRA_VencDebito = (SELECT MIN(TRAN_Vencimento) FROM Transacoes WITH (NOLOCK) WHERE CTRA_ID = " & CTRA_ID_NEW & ") WHERE CTRA_ID = " & CTRA_ID_NEW, FilialDestino, CONT_ID

			'**********************************************************************************************************************
			' Log de Andamento
			'**********************************************************************************************************************
			Andamento = "Contrato Criado Automaticamente. Motivo: Contrato original com acordo por título."
			'SQLScript = "EXEC prc_CriaAndamentoCobranca @CTRA_ID = " & CTRA_ID_NEW & ", @ANCO_DataHora = '" & now & "', @STAC_ID = 66, @TDEV_ID = NULL, @ANCO_Descricao = '" & Andamento & "' , @FUNC_ID = " & Session("FUNC_ID")
			'BD.Execute SQLScript
			'PreencheLOG_Contratante SQLScript , FilialDestino, CONT_ID
			
			'**********************************************************************************************************************
			' Verifica se o contrato novo só tem título a vencer, caso positivo, devolve o contrato
			'**********************************************************************************************************************
			Set rsT2 = BD.Execute("SELECT TRAN_ID FROM Transacoes WITH (NOLOCK) WHERE CTRA_ID = " & CTRA_ID_NEW & " AND TRAN_Vencimento < getdate() ")
			'if rsT2.EOF then
				'Gera bordero de Prestação de Contas
				'vID = RetornaProximoNumero("Bordero", "BORD_ID", Session("FILI_ID"))
				'SQLScript = "INSERT INTO Bordero (BORD_ID,CART_ID,TBOR_ID,BORD_DataEmissao,BORD_DataDevolucao,BORD_PzExclusividade) VALUES(" & vID &  "," & Carteira & ",'D','" & date & "','" & date & "',0)"
				'BD.Execute SQLScript
				''PreencheLOG_Contratante SQLScript, FilialDestino, CONT_ID

				'Relaciona Contrato com Borderô
				'SQLScript = "EXEC prc_CriaRelContratoBordero2 " & CTRA_ID_NEW & ", " & vID & ", 98, '" & date & "', 1"
				'BD.Execute SQLScript
				''PreencheLOG_Contratante SQLScript, FilialDestino, CONT_ID
				
				'Atualiza Contrato
				'SQLScript = "UPDATE Contratos SET SCON_ID = 4 WHERE CTRA_ID = " & CTRA_ID_NEW
				'BD.Execute SQLScript
				''PreencheLOG_Contratante SQLScript, FilialDestino, CONT_ID

				'Insere Andamento de Cobrança
				'SQLScript = "EXEC prc_CriaAndamentoCobranca @CTRA_ID = " & CTRA_ID_NEW & ", @ANCO_DataHora = '" & now & "', @STAC_ID = 1, @TDEV_ID = NULL, @ANCO_Descricao = 'Devolução. Motivo: Contrato sem títulos vencidos.', @FUNC_ID = " & Session("FUNC_ID")
				'BD.Execute SQLScript
				''PreencheLOG_Contratante SQLScript, FilialDestino, CONT_ID
			'end if
		end if
	end if

End Sub

Function ValorParcelaComJurosParcelamentoHSBC(Principal, Juros, QtdParcelas, DataAcordo, DataEntrada, ValorAcordo, ValorIOF)
	'Principal: 1000,00
	'Juros: 0,02 (para juros de 2%)
	'QtdParcelas: 10
	'QtdCasasDecimeis: Se for -1, significa que deve-se levar em conta todas as casas decimais
	'ValorAcordo: Somatório das origens atualizada até a data da entrada
	Dim fvIndiceDiario, i
	Dim frsValidaData, fvDataParcela, fvPrazoAcumuladoParcela, fvIndiceCorrecaoParcela, fvFatorAcumulado
	Dim fvValorParcela, fvSaldoDevedor, fvJurosParcela, fvPrincipalParcela, fvDataParcelaAnterior
	Dim fvPrazoParcela, fvIndiceIOFParcela, fvValorIOFAcumulado, fvValorIOFParcela, fvFatorParcela
	Dim fvIndiceIOFContrato, fvValorIOFContrato, fvValorParcelaFinal, fvDataBase
	
	fvIndiceDiario = Round(((1 + Juros)^(1/30)), 7)

	fvFatorAcumulado = 0
	For i = 1 to QtdParcelas
		fvDataParcela = DateAdd("m", i, DataEntrada)
		Set frsValidaData = BD.Execute("SELECT dbo.func_VerificaData(" & Session("FILI_ID") & ", '" & fvDataParcela & "') Data")
		if not frsValidaData("Data") then
			Set frsValidaData = BD.Execute("SELECT dbo.func_ValidaDiaUtil(" & Session("FILI_ID") & ", '" & fvDataParcela & "') Data")
			'Set frsValidaData = BD.Execute("SELECT dbo.func_CalculaDataUtilFutura(" & Session("FILI_ID") & ", '" & fvDataParcela & "', 1) Data")
			fvDataParcela = frsValidaData("Data")
		end if
		fvPrazoAcumuladoParcela = DateDiff("d", DataEntrada, fvDataParcela)
		fvIndiceCorrecaoParcela = CDbl(Mid(CStr(fvIndiceDiario ^ fvPrazoAcumuladoParcela), 1, 9))
		fvFatorParcela = TruncaValor(1 / fvIndiceCorrecaoParcela, 7)
		fvFatorAcumulado = fvFatorAcumulado + fvFatorParcela
	Next
	fvValorParcela = TruncaValor(Principal / fvFatorAcumulado, 2)

	fvSaldoDevedor = Principal
	fvDataParcela = DateAdd("m", 1, DataEntrada)
	Set frsValidaData = BD.Execute("SELECT dbo.func_VerificaData(" & Session("FILI_ID") & ", '" & fvDataParcela & "') Data")
	if not frsValidaData("Data") then
		Set frsValidaData = BD.Execute("SELECT dbo.func_ValidaDiaUtil(" & Session("FILI_ID") & ", '" & fvDataParcela & "') Data")
		'Set frsValidaData = BD.Execute("SELECT dbo.func_CalculaDataUtilFutura(" & Session("FILI_ID") & ", '" & fvDataParcela & "', 1) Data")
		fvDataParcela = frsValidaData("Data")
	end if
	fvDataParcelaAnterior = DataEntrada
	For i = 1 to QtdParcelas
		fvPrazoParcela = DateDiff("d", fvDataParcelaAnterior, fvDataParcela)
		fvJurosParcela = TruncaValor(fvSaldoDevedor * ((fvIndiceDiario ^ fvPrazoParcela) - 1), 2)
		fvPrincipalParcela = fvValorParcela - fvJurosParcela
		fvSaldoDevedor = fvSaldoDevedor - fvPrincipalParcela
		fvPrazoAcumuladoParcela = DateDiff("d", DataEntrada, fvDataParcela)
		if fvPrazoAcumuladoParcela <= 365 then
			fvIndiceIOFParcela = TruncaValor(0.000041 * fvPrazoAcumuladoParcela, 6)
		else
			fvIndiceIOFParcela = 0.014965
		end if
		fvValorIOFParcela = TruncaValor(fvPrincipalParcela * fvIndiceIOFParcela, 2)
		'Response.Write fvPrincipalParcela & " - " & fvDataParcela & " - " & fvIndiceIOFParcela & " - " & fvPrazoAcumuladoParcela & " - " & fvValorIOFParcela & "<br>"
		fvValorIOFAcumulado = fvValorIOFAcumulado + fvValorIOFParcela
		fvDataParcelaAnterior = fvDataParcela
		fvDataParcela = DateAdd("m", i + 1, DataEntrada)
		Set frsValidaData = BD.Execute("SELECT dbo.func_VerificaData(" & Session("FILI_ID") & ", '" & fvDataParcela & "') Data")
		if not frsValidaData("Data") then
			Set frsValidaData = BD.Execute("SELECT dbo.func_ValidaDiaUtil(" & Session("FILI_ID") & ", '" & fvDataParcela & "') Data")
			'Set frsValidaData = BD.Execute("SELECT dbo.func_CalculaDataUtilFutura(" & Session("FILI_ID") & ", '" & fvDataParcela & "', 1) Data")
			fvDataParcela = frsValidaData("Data")
		end if
	Next
	fvIndiceIOFContrato = TruncaValor((fvValorIOFAcumulado / ValorAcordo) / (1 - (fvValorIOFAcumulado / ValorAcordo)), 7)
	fvValorIOFContrato = TruncaValor(ValorAcordo * fvIndiceIOFContrato, 2)
	fvSaldoDevedor = Principal + fvValorIOFContrato
	fvValorParcelaFinal = TruncaValor((fvSaldoDevedor / fvFatorAcumulado) * 1.01, 2)
	
	ValorIOF = fvValorIOFContrato
	
	ValorParcelaComJurosParcelamentoHSBC = fvValorParcelaFinal
End Function

Function TruncaValor(Valor, CasasDecimais)
	Dim vValorArredondamento
	vValorArredondamento = CDbl("0," & String(CasasDecimais - 1, "0") & "1")
	if Round(Valor, CasasDecimais) > Valor then
		TruncaValor = Round(Valor, CasasDecimais) - vValorArredondamento
	else
		TruncaValor = Round(Valor, CasasDecimais)
	end if
End Function

Function ValorParcelaComJurosParcelamento(Principal, Juros, QtdParcelas, QtdCasasDecimeis)
	'Principal: 1000,00
	'Juros: 0,02 (para juros de 2%)
	'QtdParcelas: 10
	'QtdCasasDecimeis: Se for -1, significa que deve-se levar em conta todas as casas decimais
	
	if QtdCasasDecimeis = "-1" or QtdCasasDecimeis = -1 then
		if Juros = 0 then
			ValorParcelaComJurosParcelamento = Principal / QtdParcelas
		else
			ValorParcelaComJurosParcelamento = Principal * ((Juros * ((1 + Juros)^QtdParcelas)) / ((1 + Juros)^QtdParcelas - 1))
		end if
	else
		ValorParcelaComJurosParcelamento = Principal * Round(((Juros * ((1 + Juros)^QtdParcelas)) / ((1 + Juros)^QtdParcelas - 1)),QtdCasasDecimeis)
		'ValorParcelaComJurosParcelamento = Principal * CDbl(Mid(CStr(Round((Juros * ((1 + Juros)^(QtdParcelas))) / ((1 + Juros)^QtdParcelas - 1),10)), 1, QtdCasasDecimeis + 2))
	end if
End Function

Function ValorParcelaComJurosParcelamentoCacique(Principal, Juros, QtdParcelas, QtdCasasDecimeis)
	'Principal: 1000,00
	'Juros: 0,02 (para juros de 2%)
	'QtdParcelas: 10
	'QtdCasasDecimeis: Se for -1, significa que deve-se levar em conta todas as casas decimais
	
	if Juros = 0 then
		ValorParcelaComJurosParcelamentoCacique = Principal / QtdParcelas
	else
		ValorParcelaComJurosParcelamentoCacique = (Principal * (1 + Juros)) / QtdParcelas
	end if
End Function

Function ValorParcelaComJurosParcelamentoAvon(Principal, Juros, QtdParcelas, QtdCasasDecimeis)
	'Principal: 1000,00
	'Juros: 0,02 (para juros de 2%)
	'QtdParcelas: 10
	'QtdCasasDecimeis: Se for -1, significa que deve-se levar em conta todas as casas decimais
	
	'if QtdCasasDecimeis = "-1" or QtdCasasDecimeis = -1 then
		if Juros = 0 then
			ValorParcelaComJurosParcelamentoAvon = Principal / QtdParcelas
		else
			ValorParcelaComJurosParcelamentoAvon = (Principal + (Principal * (Juros * QtdParcelas))) / QtdParcelas
		end if
	'else
	'	ValorParcelaComJurosParcelamentoAvon = Round((Principal + (Principal * (Juros * QtdParcelas))) / QtdParcelas, QtdCasasDecimeis)
	'end if
End Function

Function ValorParcelaComJurosParcelamentoHiper(Principal, Juros, QtdParcelas, QtdCasasDecimeis)
	'Principal: 1000,00
	'Juros: 0,02 (para juros de 2%)
	'QtdParcelas: 10
	'QtdCasasDecimeis: Se for -1, significa que deve-se levar em conta todas as casas decimais
	
	if QtdCasasDecimeis = "-1" or QtdCasasDecimeis = -1 then
		ValorParcelaComJurosParcelamentoHiper = Principal * ((Juros * ((1 + Juros)^QtdParcelas)) / ((1 + Juros)^QtdParcelas - 1))
	else
		ValorParcelaComJurosParcelamentoHiper = Principal * Round(((Juros * ((1 + Juros)^QtdParcelas)) / ((1 + Juros)^QtdParcelas - 1)),QtdCasasDecimeis)
		'ValorParcelaComJurosParcelamento = Principal * CDbl(Mid(CStr(Round((Juros * ((1 + Juros)^(QtdParcelas))) / ((1 + Juros)^QtdParcelas - 1),10)), 1, QtdCasasDecimeis + 2))
	end if
End Function

Function FatorParcelamento(Juros, QtdParcelas, QtdCasasDecimeis)
	'Principal: 1000,00
	'Juros: 0,02 (para juros de 2%)
	'QtdParcelas: 10
	'QtdCasasDecimeis: Se for -1, significa que deve-se levar em conta todas as casas decimais

	if QtdCasasDecimeis = "-1" or QtdCasasDecimeis = -1 then
		FatorParcelamento = (Juros * ((1 + Juros)^(QtdParcelas))) / ((1 + Juros)^QtdParcelas - 1)
	else
		FatorParcelamento = Round((Juros * ((1 + Juros)^(QtdParcelas))) / ((1 + Juros)^QtdParcelas - 1),QtdCasasDecimeis)
		'FatorParcelamento = CDbl(Mid(CStr(Round((Juros * ((1 + Juros)^(QtdParcelas))) / ((1 + Juros)^QtdParcelas - 1),10)), 1, QtdCasasDecimeis + 2))
		'FatorParcelamento = Round((Juros * ((1 + Juros)^(QtdParcelas))) / ((1 + Juros)^QtdParcelas - 1),10)
	end if
End Function

Function ValorPrincipal(Contratante, Valor, DataVencDebito, VencTransac, Produto)
	Dim vValorAtualizado, vValorDebito, vValorMulta, vValorEncargos
	'if Contratante = 7 and (DateDiff("d", DataVencDebito, date) > 180 or DateDiff("d", DataVencDebito, date) < -180) and VencTransac < date and Valor < 5000 and Produto = 301 then
	'	vValorAtualizado = 0
	'	vValorDebito = Valor
	'	vValorMulta = 0
	'	vValorEncargos = 0
	'	
	'	vValorEncargos = vValorDebito * 0.105
	'	vValorMulta = vValorDebito * 0.02
	'	vValorAtualizado = vValorDebito + vValorEncargos + vValorMulta
	'
	'	vValorEncargos = Round(vValorAtualizado * 0.105, 2)
	'	vValorMulta = Round(vValorAtualizado * 0.02, 2)
	'	vValorAtualizado = vValorAtualizado + vValorEncargos + vValorMulta
	'
	'	ValorPrincipal = vValorAtualizado
	'else
		ValorPrincipal = Valor
	'end if
End Function

'Formata o contrato. Ex.: 0000.0000.0000.0000 
function FormataContrato(valor) 
	Dim vr, tam
	vr = valor
	tam = len(vr)

	if tam <= 4 then FormataContrato = vr
	if tam >= 5 and tam <= 8 then FormataContrato = mid(vr, 1, 4) & "." & mid(vr, 5, 4)
	if tam >= 9 and tam <= 12 then FormataContrato = mid(vr, 1, 4) & "." & mid(vr, 5, 4) & "." & mid(vr, 9, 4)
	if tam >= 13 and tam <= 16 then FormataContrato = mid(vr, 1, 4) & "." & mid(vr, 5, 4) & "." & mid(vr, 9, 4) & "." & mid(vr, 13, 4)
	if tam >= 17 and tam <= 20  then FormataContrato = mid(vr, 1, 4) & "." & mid(vr, 5, 4) & "." & mid(vr, 9, 4) & "." & mid(vr, 13, 4) & "." & mid(vr, 17, 4)
End Function

'Formata o CPF e o CNPJ, aparecendo com pontos, e barras. Ex.: 000.000.000-00 ou 000.000.000/0000-00 
function FormataCPFCNPJ(CPF_CNPJ)
	if len(CPF_CNPJ) = 11 then
		FormataCPFCNPJ = mid(CPF_CNPJ,1,3) & "." & mid(CPF_CNPJ,4,3) & "." & mid(CPF_CNPJ,7,3) & "-" & mid(CPF_CNPJ,10,2)
	else
		FormataCPFCNPJ = mid(CPF_CNPJ,1,2) & "." & mid(CPF_CNPJ,3,3) & "." & mid(CPF_CNPJ,6,3) & "/" & mid(CPF_CNPJ,9,4) & "-" & mid(CPF_CNPJ,13,2)
	end if
end function

'Formata o CEP. Ex.: 00.000-000 
function FormataCEP(CEP)
		FormataCEP = mid(CEP,1,2) & mid(CEP,3,3) & "-" & mid(CEP,6,3)
end function

'Atualiza o valor de um título baseado na Politica de Negociação da empresa
function AtualizaValor(ByVal ID_Carteira, ByVal VencimentoDebito, ByVal AtualizaAte, ByVal Valor, ByVal Sinal, ByVal PermiteAtualizacao, ByVal QtdParc, ByVal DataRecebimento, ByVal TDOC_ID, ByVal DataAtualizacao, ByVal DataVencDebito)
	Dim Conn
	Dim RSIGPMHoje
	Dim RSIGPMVencimento
	Dim RSPolitica
	Dim vAtualiza, vAtualizaDebito, vAtualizaCredito, vMulta, vPercentMulta, vMultaPrincipal, vMultaPrincipalCorrigido, vJuros, vTaxaJuros, vPercentJuros
	Dim vNumDias, vValorCorrigido, vValorMulta, vValorHonorarios, vHonorarios, vPercentHonorarios, vHonorariosPrincipal, vHonorariosPrincipalCorrigido
	Dim vJurosSimples, vJurosComposto, vJurosFaixa, RSTaxaJurosFaixa
	Dim rsCont, vQtdMeses, sai, vDataNova, vValorJuros
	Dim vVencimentoTitulo, vRecebimentoContrato, vAtualizaValor, vTaxaAtualizacao, vNumDias2, vNumDias3
	Dim vTaxaJurosComposta
	Dim vValorMultaIOF, vValorMultaObrig, vValorIOFObrig
	Dim vCONT_ID
	Dim vAtraso
	Dim vNumeroCarteira
	
	vQtdMeses = 0
	vValorJuros = 0
	vValorMulta = 0
	vValorCorrigido = Valor
	sai = false
	
	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open Application("Conexao")

	Set RSPolitica = Conn.Execute("SELECT * FROM Politica_de_Negociacao p WITH (NOLOCK) JOIN REL_PolNeg_Cart r ON p.PNEG_ID = r.PNEG_ID WHERE CART_ID = " & ID_Carteira)
	Set rsCont = Conn.Execute("SELECT CONT_ID, CART_Numero FROM Carteiras WITH (NOLOCK) WHERE CART_ID = " & ID_Carteira)	
	vCONT_ID = rsCont("CONT_ID")
	vNumeroCarteira = rsCont("CART_Numero")
	
	if not RSPolitica.eof then
		vHonorarios = RSPolitica("PNEG_Honorarios")
		vPercentHonorarios = RSPolitica("PNEG_PercentHonorarios")
		vHonorariosPrincipal = RSPolitica("PNEG_HonorariosPrincipal")
		vHonorariosPrincipalCorrigido = RSPolitica("PNEG_HonorariosPrincipalCorrigido")

		if vCONT_ID = 101 then 'ID_Carteira = 18 or ID_Carteira = 202 then
			vAtraso = DateDiff("d", DataVencDebito, AtualizaAte)
			if vAtraso <= 150 then
				vPercentHonorarios = 10
			elseif vAtraso >= 151 and vAtraso <= 360 and QtdParc = 1 then
				vPercentHonorarios = 15
			elseif vAtraso >= 151 and vAtraso <= 360 and QtdParc > 1 then
				vPercentHonorarios = 10
			end if
		'elseif vCONT_ID = 34 or vCONT_ID = 100 then
		'	vAtraso = DateDiff("d", DataVencDebito, AtualizaAte)
		'	if vAtraso <= 60 then
		'		vPercentHonorarios = 0
		'	end if
		end if
	end if
	
	Conn.Close 
	Set Conn = Nothing

	if PermiteAtualizacao and VencimentoDebito < AtualizaAte then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open Application("Conexao")
	
		Set RSPolitica = Conn.Execute("SELECT * FROM Politica_de_Negociacao p WITH (NOLOCK) JOIN REL_PolNeg_Cart r ON p.PNEG_ID = r.PNEG_ID WHERE CART_ID = " & ID_Carteira)
		Set rsCont = Conn.Execute("SELECT CONT_ID FROM Carteiras WITH (NOLOCK) WHERE CART_ID = " & ID_Carteira)	
		
		if not RSPolitica.eof then
			vAtualiza = RSPolitica("PNEG_Atualiza")
			vAtualizaDebito = RSPolitica("PNEG_AtualizaDebito")
			vAtualizaCredito = RSPolitica("PNEG_AtualizaCredito")
			vMulta = RSPolitica("PNEG_Multa")
			vPercentMulta = RSPolitica("PNEG_PercentMulta")
			vMultaPrincipal = RSPolitica("PNEG_MultaPrincipal")
			vMultaPrincipalCorrigido = RSPolitica("PNEG_MultaPrincipalCorrigido")
			vAtualizaValor = RSPolitica("PNEG_AtualizaValor")
			vTaxaAtualizacao = RSPolitica("TAXA_ID_Atualizacao")
			vJuros = RSPolitica("PNEG_Juros")
			vJurosFaixa = RSPolitica("PNEG_JurosPorFaixa")
			vPercentJuros = RSPolitica("PNEG_PercentualJuros")
			vJurosSimples = RSPolitica("PNEG_JurosSimples")
			vJurosComposto = RSPolitica("PNEG_JurosComposto")
			vHonorarios = RSPolitica("PNEG_Honorarios")
			vPercentHonorarios = RSPolitica("PNEG_PercentHonorarios")
			vHonorariosPrincipal = RSPolitica("PNEG_HonorariosPrincipal")
			vHonorariosPrincipalCorrigido = RSPolitica("PNEG_HonorariosPrincipalCorrigido")
			vVencimentoTitulo = RSPolitica("PNEG_AtualizaDoVencimento")
			vRecebimentoContrato = RSPolitica("PNEG_AtualizaDoRecebimento")
			
			if vCONT_ID = 101 then 'ID_Carteira = 18 or ID_Carteira = 202 then
				vAtraso = DateDiff("d", DataVencDebito, AtualizaAte)
				if vAtraso <= 150 then
					vPercentHonorarios = 10
				elseif vAtraso >= 151 and vAtraso <= 360 and QtdParc = 1 then
					vPercentHonorarios = 15
				elseif vAtraso >= 151 and vAtraso <= 360 and QtdParc > 1 then
					vPercentHonorarios = 10
				end if
			'elseif vCONT_ID = 34 or vCONT_ID = 100 then
			'	vAtraso = DateDiff("d", DataVencDebito, AtualizaAte)
			'	if vAtraso <= 60 then
			'		vPercentHonorarios = 0
			'	end if
			end if
			
			if (ID_Carteira = 41 or ID_Carteira = 42 or ID_Carteira = 43 or ID_Carteira = 44) and AtualizaAte > CDate("17/10/2005") then
				AtualizaAte = CDate("17/10/2005")
			end if
			if vCONT_ID = 43 and AtualizaAte > CDate("10/04/2006") then
				AtualizaAte = CDate("10/04/2006")
			end if
			
			if vRecebimentoContrato then
				vNumDias = DateDiff("d", DataRecebimento, AtualizaAte)
			else
				vNumDias = DateDiff("d", VencimentoDebito, AtualizaAte)
			end if
			
			vNumDias2 = DateDiff("d", DataVencDebito, AtualizaAte)

			Dim rsTaxaHoje, rsTaxaVenc
			Dim vNumMeses, vDiasRestantes, w, rsDescTipoDoc, vDescTipoDoc
			
			vDescTipoDoc = ""
			if rsCont("CONT_ID") = 10 or rsCont("CONT_ID") = 31 then
				Set rsDescTipoDoc = BD.Execute("SELECT TDOC_Descricao FROM Tipos_de_Documento WITH (NOLOCK) WHERE TDOC_ID = " & TDOC_ID)
				if not rsDescTipoDoc.EOF then
					vDescTipoDoc = rsDescTipoDoc("TDOC_Descricao")
				end if
				rsDescTipoDoc.Close
				Set rsDescTipoDoc =  nothing
			end if
			
			if rsCont("CONT_ID") = 83 then
				vNumDias3 = vNumDias
				if vNumDias3 > 114 then
					vNumDias3 = 114
				end if
				vValorJuros = (vValorCorrigido * (1 + (12.9 / 100))^(vNumDias3 / 30)) - vValorCorrigido
				vValorCorrigido = vValorCorrigido + vValorJuros
			elseif rsCont("CONT_ID") = 10 then
				if vDescTipoDoc = "ADIANTAMENTO DEPOSITANTE" or vDescTipoDoc = "CHEQUE ESPECIAL" or vDescTipoDoc = "CHEQUE UNIVERSITARIO MB" or vDescTipoDoc = "CHEQUE EMPRESA MB" then
					if vNumDias < 180 then
						vPercentJuros = 7
					else
						vPercentJuros = 8
					end if
				else
					if vNumDias < 180 then
						vPercentJuros = 6
					else
						vPercentJuros = 7
					end if
				end if
				vValorJuros = (vValorCorrigido * (1 + (vPercentJuros / 100))^(vNumDias / 30)) - vValorCorrigido
				vValorCorrigido = vValorCorrigido + vValorJuros
			elseif rsCont("CONT_ID") = 99 then
				vValorCorrigido = Round(Valor * (1.0002125^vNumDias), 2)
				vValorMulta = Round(vValorCorrigido * (vPercentMulta / 100), 2)
				vValorJuros = Round((Valor + vValorMulta) * ((vNumDias * (vPercentJuros / 30)) / 100), 2)
				vValorCorrigido = vValorCorrigido + vValorJuros + vValorMulta
			elseif rsCont("CONT_ID") = 50 then
				if vNumDias < 180 then
					vValorMulta = vValorCorrigido * 0.02
					vNumMeses = vNumDias \ 30
					vDiasRestantes = vNumDias mod 30
					if vNumMeses > 0 then
						For w = 1 to vNumMeses
							vValorCorrigido = vValorCorrigido * 1.099 ' 9,9% ao mês
						Next
					end if
					if vDiasRestantes > 0 then
						vValorCorrigido = vValorCorrigido * (1 + (9.9 / 3000 * vDiasRestantes))
					end if
					vValorCorrigido = vValorCorrigido + vValorMulta
				else
					Set rsTaxaVenc = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & "01/" & Month(VencimentoDebito) & "/" & Year(VencimentoDebito) & "') AND TAXA_ID = 10")
				
					if rsTaxaVenc.EOF then
						Set rsTaxaVenc = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & "01/" & Month(VencimentoDebito) & "/" & Year(VencimentoDebito) & "') AND TAXA_ID = 10 ORDER BY COTA_Data DESC")
					end if
					
					if not rsTaxaVenc.EOF then
						vValorCorrigido = vValorCorrigido * (1 + rsTaxaVenc("COTA_Indice"))
					end if

					rsTaxaVenc.Close
				end if
			elseif TDOC_ID = 67 and DateDiff("d", VencimentoDebito, AtualizaAte) < 61 then
				'Mora
				Set rsTaxaHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 4")
				Set rsTaxaVenc = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 4")

				if rsTaxaHoje.EOF then
					Set rsTaxaHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 4 ORDER BY COTA_Data DESC")
				end if
				
				if rsTaxaVenc.EOF then
					Set rsTaxaVenc = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 4 ORDER BY COTA_Data DESC")
				end if
				
				if not rsTaxaHoje.EOF and not rsTaxaVenc.EOF then
					vValorCorrigido = (rsTaxaHoje("COTA_Indice") / rsTaxaVenc("COTA_Indice")) * vValorCorrigido
				end if

				rsTaxaVenc.Close
				rsTaxaHoje.Close
			elseif TDOC_ID = 67 and DateDiff("d", VencimentoDebito, AtualizaAte) > 60 and DateDiff("d", VencimentoDebito, DataAtualizacao) < 61 then
				'Mora
				Set rsTaxaHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & DateAdd("d", 60, VencimentoDebito) & "') AND TAXA_ID = 4")
				Set rsTaxaVenc = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 4")

				if rsTaxaHoje.EOF then
					Set rsTaxaHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & DateAdd("d", 60, VencimentoDebito) & "') AND TAXA_ID = 4 ORDER BY COTA_Data DESC")
				end if
				
				if rsTaxaVenc.EOF then
					Set rsTaxaVenc = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 4 ORDER BY COTA_Data DESC")
				end if
				
				if not rsTaxaHoje.EOF and not rsTaxaVenc.EOF then
					vValorCorrigido = (rsTaxaHoje("COTA_Indice") / rsTaxaVenc("COTA_Indice")) * vValorCorrigido
				end if

				rsTaxaVenc.Close
				rsTaxaHoje.Close

				'CL/LP
				Set rsTaxaHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 5")
				Set rsTaxaVenc = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & DateAdd("d", 60, VencimentoDebito) & "') AND TAXA_ID = 5")

				if rsTaxaHoje.EOF then
					Set rsTaxaHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 5 ORDER BY COTA_Data DESC")
				end if
				
				if rsTaxaVenc.EOF then
					Set rsTaxaVenc = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & DateAdd("d", 60, VencimentoDebito) & "') AND TAXA_ID = 5 ORDER BY COTA_Data DESC")
				end if
				
				if not rsTaxaHoje.EOF and not rsTaxaVenc.EOF then
					vValorCorrigido = (rsTaxaHoje("COTA_Indice") / rsTaxaVenc("COTA_Indice")) * vValorCorrigido
				end if

				rsTaxaVenc.Close
				rsTaxaHoje.Close
				Set rsTaxaVenc = Nothing
				Set rsTaxaHoje = Nothing
				Set rsTaxaVenc = Nothing
				Set rsTaxaHoje = Nothing
			elseif (TDOC_ID = 67 and DateDiff("d", VencimentoDebito, AtualizaAte) > 60 and DateDiff("d", VencimentoDebito, DataAtualizacao) > 60) or TDOC_ID = 68 or TDOC_ID = 69 then
				'CL/LP
				Set rsTaxaHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 5")
				Set rsTaxaVenc = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 5")

				if rsTaxaHoje.EOF then
					Set rsTaxaHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 5 ORDER BY COTA_Data DESC")
				end if
				
				if rsTaxaVenc.EOF then
					Set rsTaxaVenc = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 5 ORDER BY COTA_Data DESC")
				end if
				
				if not rsTaxaHoje.EOF and not rsTaxaVenc.EOF then
					vValorCorrigido = (rsTaxaHoje("COTA_Indice") / rsTaxaVenc("COTA_Indice")) * vValorCorrigido
				end if

				rsTaxaVenc.Close
				rsTaxaHoje.Close
				Set rsTaxaVenc = Nothing
				Set rsTaxaHoje = Nothing
			else
				if vAtualiza then
					if (vAtualizaDebito and Sinal = "+") or (vAtualizaCredito and Sinal = "-") then
						if vJuros then
							vValorJuros = 0
							'Cobrar juros
							if vJurosFaixa then
								'Juros por faixa de atraso
								'Set RSTaxaJurosFaixa = Conn.Execute("SELECT FNEG_PercentJuros, FNEG_JurosSimples, FNEG_JurosComposto FROM Faixas_de_Negociacao f WITH (NOLOCK) JOIN Tabelas_de_Negociacao t WITH (NOLOCK) ON f.TNEG_ID = t.TNEG_ID JOIN REL_TabNeg_Cart r WITH (NOLOCK) ON t.TNEG_ID = r.TNEG_ID WHERE " & vNumDias2 & " BETWEEN TNEG_InicioFaixaAtraso AND TNEG_FinalFaixaAtraso AND CART_ID = " & ID_Carteira & " AND FNEG_QtdParcelas = " & QtdParc & " AND FNEG_PercentJuros IS NOT NULL AND FNEG_Habilitado = 1")
								Set RSTaxaJurosFaixa = Conn.Execute("SELECT FNEG_PercentJuros, FNEG_JurosSimples, FNEG_JurosComposto FROM Faixas_de_Negociacao f WITH (NOLOCK) JOIN Tabelas_de_Negociacao t WITH (NOLOCK) ON f.TNEG_ID = t.TNEG_ID JOIN REL_TabNeg_Cart r WITH (NOLOCK) ON t.TNEG_ID = r.TNEG_ID WHERE ((" & vNumDias2 & " BETWEEN TNEG_InicioFaixaAtraso AND TNEG_FinalFaixaAtraso) OR (" & Year(DataVencDebito) & " BETWEEN TNEG_InicioAnoAtraso AND TNEG_FinalAnoAtraso)) AND CART_ID = " & ID_Carteira & " AND FNEG_QtdParcelas = " & QtdParc & " AND FNEG_PercentJuros IS NOT NULL AND FNEG_Habilitado = 1 AND TNEG_Habilitada = 1 AND TNEG_VigenciaDe <= '" & AtualizaAte & "' AND TNEG_VigenciaAte >= '" & AtualizaAte & "'")
								if not RSTaxaJurosFaixa.EOF then
									vJurosSimples = RSTaxaJurosFaixa("FNEG_JurosSimples")
									vJurosComposto = RSTaxaJurosFaixa("FNEG_JurosComposto")
									vPercentJuros = RSTaxaJurosFaixa("FNEG_PercentJuros")

									if vJurosSimples then
										if rsCont("CONT_ID") = 31 then
											vValorJuros = (((vNumDias * Round(vPercentJuros / 30, 3)) / 100) * vValorCorrigido)
										else
											vValorJuros = (((vNumDias * (vPercentJuros / 30)) / 100) * vValorCorrigido)
										end if
									elseif vJurosComposto then
										vValorJuros = (vValorCorrigido * (1 + (vPercentJuros / 100))^(vNumDias / 30)) - vValorCorrigido
									end if
								end if							
								RSTaxaJurosFaixa.Close
								Set RSTaxaJurosFaixa = Nothing
							elseif vPercentJuros <> "" then
								if vJurosSimples then
									if rsCont("CONT_ID") = 31 then
										vValorJuros = (((vNumDias * Round(vPercentJuros / 30, 3)) / 100) * vValorCorrigido)
									else
										vValorJuros = (((vNumDias * (vPercentJuros / 30)) / 100) * vValorCorrigido)
									end if
								elseif vJurosComposto then
									vValorJuros = (vValorCorrigido * (1 + (vPercentJuros / 100))^(vNumDias / 30)) - vValorCorrigido
								end if
							end if
						end if
						if vAtualizaValor and UCASE(vDescTipoDoc) <> "CREDUCSAL" then
							if vTaxaAtualizacao = 3 then
								'IGPM
								Set RSIGPMHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 3")
								Set RSIGPMVencimento = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & VencimentoDebito & "') + 2 AND TAXA_ID = 3")

								if RSIGPMHoje.EOF then
									Set RSIGPMHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 3 ORDER BY COTA_Data DESC")
								end if
								
								if RSIGPMVencimento.EOF then
									Set RSIGPMVencimento = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & VencimentoDebito & "') AND TAXA_ID = 3 ORDER BY COTA_Data DESC")
								end if
								
								if RSIGPMHoje.EOF then
									Set RSIGPMHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE TAXA_ID = 3 ORDER BY COTA_Data DESC")
								end if
								
								if RSIGPMVencimento.EOF then
									Set RSIGPMVencimento = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE TAXA_ID = 3 ORDER BY COTA_Data DESC")
								end if

								vValorCorrigido = (RSIGPMHoje("COTA_Indice") / RSIGPMVencimento("COTA_Indice")) * vValorCorrigido
								
								RSIGPMVencimento.Close
								RSIGPMHoje.Close
								Set RSIGPMVencimento = Nothing
								Set RSIGPMHoje = Nothing
							elseif vTaxaAtualizacao = 10 then
								'Taxa Renner
								Set RSIGPMHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = " & vTaxaAtualizacao)
								Set RSIGPMVencimento = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & VencimentoDebito & "') AND TAXA_ID = " & vTaxaAtualizacao)

								if RSIGPMHoje.EOF then
									Set RSIGPMHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
								end if
								
								if RSIGPMVencimento.EOF then
									Set RSIGPMVencimento = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & VencimentoDebito & "') AND TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
									if RSIGPMVencimento.EOF then
										Set RSIGPMVencimento = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data")
									end if
								end if
								
								'if RSIGPMHoje.EOF or RSIGPMVencimento.EOF then
								'	vValorCorrigido = Valor
								'else
									vValorCorrigido = (RSIGPMVencimento("COTA_Indice") / RSIGPMHoje("COTA_Indice") ) * vValorCorrigido
								'end if
								
								
								RSIGPMVencimento.Close
								RSIGPMHoje.Close
								Set RSIGPMVencimento = Nothing
								Set RSIGPMHoje = Nothing
							elseif vTaxaAtualizacao = 15 then
								'IGPM Pró-rata
								if CDate(AtualizaAte) > CDate(VencimentoDebito) then
									if Month(VencimentoDebito) = Month(AtualizaAte) and Year(VencimentoDebito) = Year(AtualizaAte) then
										Set RSIGPMVencimento = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & "01/" & Month(VencimentoDebito) & "/" & Year(VencimentoDebito) & "') AND TAXA_ID = " & vTaxaAtualizacao)
										if not RSIGPMVencimento.EOF then
											vValorCorrigido = DateDiff("d", VencimentoDebito, AtualizaAte) * RSIGPMVencimento("COTA_Indice") / 100 / 30 * vValorCorrigido
										end if
									else
										sai = false
										vTaxaJurosComposta = 0
										Set RSIGPMVencimento = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & "01/" & Month(VencimentoDebito) & "/" & Year(VencimentoDebito) & "') AND TAXA_ID = " & vTaxaAtualizacao)
										if not RSIGPMVencimento.EOF then
											vTaxaJurosComposta = 1 + Round((DateDiff("d", VencimentoDebito, DateAdd("m", 1, ("01/" & Month(VencimentoDebito) & "/" & Year(VencimentoDebito)))) * RSIGPMVencimento("COTA_Indice") / 100 / 30), 6)
										end if
										w = VencimentoDebito
										Do While not sai
											w = DateAdd("m", 1, w)
											if DateDiff("m", w, AtualizaAte) = 0 then
												sai = true
												Set RSIGPMVencimento = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & "01/" & Month(AtualizaAte) & "/" & Year(AtualizaAte) & "') AND TAXA_ID = " & vTaxaAtualizacao)
												if not RSIGPMVencimento.EOF then
													vTaxaJurosComposta = Round(vTaxaJurosComposta * Round((1 + (DateDiff("d", "01/" & Month(AtualizaAte) & "/" & Year(AtualizaAte), AtualizaAte) * RSIGPMVencimento("COTA_Indice") / 100 / 30)), 6), 6)
												end if
											else
												Set RSIGPMVencimento = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & "01/" & Month(w) & "/" & Year(w) & "') AND TAXA_ID = " & vTaxaAtualizacao)
												if not RSIGPMVencimento.EOF then
													vTaxaJurosComposta = Round(vTaxaJurosComposta * Round((1 + RSIGPMVencimento("COTA_Indice") / 100), 6), 6)
												end if
											end if
										Loop
										vValorCorrigido = vValorCorrigido + Round((vTaxaJurosComposta - 1) * vValorCorrigido, 2)
									end if
								end if
							else
								'Qualquer outro índice
								Set RSIGPMHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = " & vTaxaAtualizacao)
								Set RSIGPMVencimento = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & VencimentoDebito & "') AND TAXA_ID = " & vTaxaAtualizacao)

								if RSIGPMHoje.EOF then
									Set RSIGPMHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
								end if
								
								if RSIGPMVencimento.EOF then
									Set RSIGPMVencimento = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & VencimentoDebito & "') AND TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
									if RSIGPMVencimento.EOF then
										Set RSIGPMVencimento = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data")
									end if
								end if
								
								vValorCorrigido = (RSIGPMHoje("COTA_Indice") / RSIGPMVencimento("COTA_Indice")) * vValorCorrigido
								
								RSIGPMVencimento.Close
								RSIGPMHoje.Close
								Set RSIGPMVencimento = Nothing
								Set RSIGPMHoje = Nothing
							end if
							if rsCont("CONT_ID") = 31 then
								vValorJuros = (((vNumDias * Round(vPercentJuros / 30, 3)) / 100) * vValorCorrigido)
							end if
						end if
						vValorCorrigido = vValorCorrigido + vValorJuros
						if vMulta then
							'Cobrar multa
							vValorMulta = 0
							if vCONT_ID = 13 and vNumDias > 120 then
								vPercentMulta = 0
							end if
							if vMultaPrincipal then
								'Cobrar multa sobre o principal
								'vValorMulta = ((vNumDias * (vPercentMulta / 30)) / 100) * Valor
								vValorMulta = (vPercentMulta / 100) * Valor
							elseif vMultaPrincipalCorrigido then
								'Cobrar multa sobre o proncipal corrigido (com Juros)
								'vValorMulta = ((vNumDias * (vPercentMulta / 30)) / 100) * vValorCorrigido
								vValorMulta = (vPercentMulta / 100) * vValorCorrigido
							end if
							if rsCont("CONT_ID") = 31 then
								if UCASE(vDescTipoDoc) = "CREDUCSAL" then
									vValorMulta = (vPercentMulta / 100) * Valor
								else
									vValorCorrigido = vValorCorrigido - vValorJuros
									vValorMulta = (vPercentMulta / 100) * vValorCorrigido
									vValorCorrigido = vValorCorrigido + vValorJuros
								end if
							end if
							vValorCorrigido = vValorCorrigido + vValorMulta
						end if
					end if
				end if
			end if
		end if
		Conn.Close 
		Set Conn = Nothing
	end if
	
	if vCONT_ID = 1 then
		vValorMultaObrig = Valor * 0.02
		vValorIOFObrig = (Valor + vValorMultaObrig) * 0.01
		vValorMultaIOF = vValorMultaObrig + vValorIOFObrig
		if vValorMultaIOF > vValorCorrigido - Valor then
			vValorCorrigido = Valor + vValorMultaIOF
		end if
	end if
	
	if vHonorarios then
		'Cobrar honorarios
		vValorHonorarios = 0
		if vHonorariosPrincipal then
			'Cobrar honorarios sobre o principal
			vValorHonorarios = (vPercentHonorarios / 100) * Valor
		elseif vHonorariosPrincipalCorrigido then
			'Cobrar honorarios sobre o proncipal corrigido (com Juros e Multa)
			vValorHonorarios = (vPercentHonorarios / 100) * vValorCorrigido
		end if
		vValorCorrigido = vValorCorrigido + vValorHonorarios
	end if

	AtualizaValor = vValorCorrigido
end function

Function ContaDiasUteis(De, Ate)
	Dim vQtdDiasUteis, vNumDias, vData, i, vDiaSemana
	vQtdDiasUteis = 0
	vNumDias = CDate(Ate) - CDate(De)
	vData = CDate(De)
	For i = 1 to vNumDias
		vData = vData + 1
		vDiaSemana = WeekDay(vData)
		if vDiaSemana <> 1 and vDiaSemana <> 7 then
			vQtdDiasUteis = vQtdDiasUteis + 1
		end if
	Next
	ContaDiasUteis = vQtdDiasUteis
End Function

'Atualiza o valor de uma parcela baseado na Politica de Negociação da empresa
function AtualizaValorParcela(ID_Carteira, VencimentoParcela, AtualizaAte, Valor, Faixa)
	Dim Conn
	Dim RSIGPMHoje
	Dim RSIGPMVencimento
	Dim RSTaxaHoje
	Dim RSTaxaVencimento
	Dim RSPolitica
	Dim vAtualiza, vAtualizaDebito, vAtualizaCredito, vMulta, vPercentMulta, vMultaPrincipal, vMultaPrincipalCorrigido, vJuros, vTaxaJuros, vPercentJuros
	Dim vNumDias, vValorCorrigido, vValorMulta, vValorHonorarios, vHonorarios, vPercentHonorarios, vHonorariosPrincipal, vHonorariosPrincipalCorrigido
	Dim vJurosSimples, vJurosComposto, vPercentFaixa, RSFaixa, rsContratante
	
	vValorCorrigido = Valor

	if ContaDiasUteis(VencimentoParcela, AtualizaAte) > 0 then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open Application("Conexao")
		
		Set RSPolitica = Conn.Execute("SELECT * FROM Politica_de_Negociacao p WITH (NOLOCK) JOIN REL_PolNeg_Cart r ON p.PNEG_ID = r.PNEG_ID WHERE CART_ID = " & ID_Carteira)
		
		Set rsContratante = Conn.Execute("SELECT CONT_ID FROM Carteiras WITH (NOLOCK) WHERE CART_ID = " & ID_Carteira)

		if not RSPolitica.eof then
			vAtualiza = RSPolitica("PNEG_CorrigeParcelaAtraso")
			vTaxaJuros = RSPolitica("TAXA_ID_CorrecaoParcela")
			vPercentJuros = RSPolitica("PNEG_PercentCorrecaoParcela")
			vPercentFaixa = RSPolitica("PNEG_CorrigeParcelaAtrasoPorFaixa")
				
			vNumDias = AtualizaAte - VencimentoParcela
			
			if vPercentFaixa then
				Set RSFaixa = BD.Execute("SELECT * FROM Faixas_de_Negociacao WITH (NOLOCK) WHERE FNEG_ID = " & Faixa)
				if Not IsNull(RSFaixa("FNEG_PercentJuros")) and RSFaixa("FNEG_PercentJuros") <> "" then
					vPercentJuros = RSFaixa("FNEG_PercentJuros")
				else
					vPercentJuros = 0
				end if
			end if
				
			if rsContratante("CONT_ID") = 7 then
				if vPercentJuros = 0 then
					vPercentJuros = 1
				end if
			end if

			if vAtualiza then
				if vTaxaJuros <> "" then
					if vTaxaJuros = 3 then
						'IGPM
						Set RSIGPMHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & date & "') AND TAXA_ID = " & vTaxaJuros)
						Set RSIGPMVencimento = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & VencimentoParcela & "') + 2 AND TAXA_ID = " & vTaxaJuros)

						if RSIGPMHoje.EOF then
							Set RSIGPMHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & date & "') AND TAXA_ID = " & vTaxaJuros & " ORDER BY COTA_Data DESC")
						end if
									
						if RSIGPMVencimento.EOF then
							Set RSIGPMVencimento = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & VencimentoParcela & "') AND TAXA_ID = " & vTaxaJuros & " ORDER BY COTA_Data DESC")
						end if
									
						'if RSIGPMHoje.EOF or RSIGPMVencimento.EOF then
						'	vValorCorrigido = Valor
						'else
							vValorCorrigido = (RSIGPMHoje("COTA_Indice") / RSIGPMVencimento("COTA_Indice")) * Valor
						'end if
									
						RSIGPMVencimento.Close
						RSIGPMHoje.Close
						Set RSIGPMVencimento = Nothing
						Set RSIGPMHoje = Nothing
					elseif vTaxaJuros = 10 then
						'Taxa Renner
						Set RSIGPMHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & date & "') AND TAXA_ID = " & vTaxaJuros)
						Set RSIGPMVencimento = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & VencimentoParcela & "') AND TAXA_ID = " & vTaxaJuros)

						if RSIGPMHoje.EOF then
							Set RSIGPMHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & date & "') AND TAXA_ID = " & vTaxaJuros & " ORDER BY COTA_Data DESC")
						end if
									
						if RSIGPMVencimento.EOF then
							Set RSIGPMVencimento = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & VencimentoParcela & "') AND TAXA_ID = " & vTaxaJuros & " ORDER BY COTA_Data DESC")
						end if
									
						'if RSIGPMHoje.EOF or RSIGPMVencimento.EOF then
						'	vValorCorrigido = Valor
						'else
							vValorCorrigido = (RSIGPMHoje("COTA_Indice") / RSIGPMVencimento("COTA_Indice")) * Valor
						'end if
									
						RSIGPMVencimento.Close
						RSIGPMHoje.Close
						Set RSIGPMVencimento = Nothing
						Set RSIGPMHoje = Nothing
					else
						'CL/LP, Mora, etc
						Set RSTaxaHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & date & "') AND TAXA_ID = " & vTaxaJuros)
						Set RSTaxaVencimento = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & VencimentoParcela & "') AND TAXA_ID = " & vTaxaJuros)

						if RSTaxaHoje.EOF then
							Set RSTaxaHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & date & "') AND TAXA_ID = " & vTaxaJuros & " ORDER BY COTA_Data DESC")
						end if
									
						if RSTaxaVencimento.EOF then
							Set RSTaxaVencimento = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & VencimentoParcela & "') AND TAXA_ID = " & vTaxaJuros & " ORDER BY COTA_Data DESC")
						end if
									
						'if RSIGPMHoje.EOF or RSIGPMVencimento.EOF then
						'	vValorCorrigido = Valor
						'else
							vValorCorrigido = (RSTaxaHoje("COTA_Indice") / RSTaxaVencimento("COTA_Indice")) * Valor
						'end if
									
						RSTaxaVencimento.Close
						RSTaxaHoje.Close
						Set RSTaxaVencimento = Nothing
						Set RSTaxaHoje = Nothing
					end if
				elseif vPercentJuros <> "" then
					if rsContratante("CONT_ID") = 44 then
						if vNumDias > 0 then
							vValorCorrigido = (((vNumDias * (vPercentJuros / 30)) / 100) * Valor) + Valor
							vValorCorrigido = vValorCorrigido + ((2 / 100) * vValorCorrigido)
						end if
					else
						if vNumDias > 0 then
							vValorCorrigido = (((vNumDias * (vPercentJuros / 30)) / 100) * Valor) + Valor
						end if
					end if
				end if
			end if
		end if

		Conn.Close 
		Set Conn = Nothing
	end if
	
	AtualizaValorParcela = vValorCorrigido

end function

function AtualizaValorParcelaBoleto(ID_Carteira, VencimentoParcela, AtualizaAte, Valor, Faixa)
	Dim Conn
	Dim RSIGPMHoje
	Dim RSIGPMVencimento
	Dim RSPolitica
	Dim vAtualiza, vAtualizaDebito, vAtualizaCredito, vMulta, vPercentMulta, vMultaPrincipal, vMultaPrincipalCorrigido, vJuros, vTaxaJuros, vPercentJuros
	Dim vNumDias, vValorCorrigido, vValorMulta, vValorHonorarios, vHonorarios, vPercentHonorarios, vHonorariosPrincipal, vHonorariosPrincipalCorrigido
	Dim vJurosSimples, vJurosComposto, vPercentFaixa, RSFaixa, rsContratante
	
	vValorCorrigido = Valor

	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open Application("Conexao")
	
	Set RSPolitica = Conn.Execute("SELECT * FROM Politica_de_Negociacao p WITH (NOLOCK) JOIN REL_PolNeg_Cart r ON p.PNEG_ID = r.PNEG_ID WHERE CART_ID = " & ID_Carteira)
	
	Set rsContratante = Conn.Execute("SELECT CONT_ID FROM Carteiras WITH (NOLOCK) WHERE CART_ID = " & ID_Carteira)
	
	if not RSPolitica.eof then
		vAtualiza = RSPolitica("PNEG_CorrigeParcelaAtraso")
		vTaxaJuros = RSPolitica("TAXA_ID_CorrecaoParcela")
		vPercentJuros = RSPolitica("PNEG_PercentCorrecaoParcela")
		vPercentFaixa = RSPolitica("PNEG_CorrigeParcelaAtrasoPorFaixa")
			
		vNumDias = AtualizaAte - VencimentoParcela
		
		if vPercentFaixa then
			Set RSFaixa = BD.Execute("SELECT * FROM Faixas_de_Negociacao WITH (NOLOCK) WHERE FNEG_ID = " & Faixa)
			if Not IsNull(RSFaixa("FNEG_PercentJuros")) and RSFaixa("FNEG_PercentJuros") <> "" then
				vPercentJuros = RSFaixa("FNEG_PercentJuros")
			else
				vPercentJuros = 0
			end if
		end if
			
		if rsContratante("CONT_ID") = 7 then
			if vPercentJuros = 0 then
				vPercentJuros = 1
			end if
		end if

		if vAtualiza then
			if vTaxaJuros <> "" then
				if vTaxaJuros = 3 then
					'IGPM
					Set RSIGPMHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & date & "') AND TAXA_ID = 3")
					Set RSIGPMVencimento = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & VencimentoParcela & "') + 2 AND TAXA_ID = 3")

					if RSIGPMHoje.EOF then
						Set RSIGPMHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & date & "') AND TAXA_ID = 3 ORDER BY COTA_Data DESC")
					end if
								
					if RSIGPMVencimento.EOF then
						Set RSIGPMVencimento = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & VencimentoParcela & "') AND TAXA_ID = 3 ORDER BY COTA_Data DESC")
					end if
								
					'if RSIGPMHoje.EOF or RSIGPMVencimento.EOF then
					'	vValorCorrigido = Valor
					'else
						vValorCorrigido = (RSIGPMHoje("COTA_Indice") / RSIGPMVencimento("COTA_Indice")) * Valor
					'end if
								
					RSIGPMVencimento.Close
					RSIGPMHoje.Close
					Set RSIGPMVencimento = Nothing
					Set RSIGPMHoje = Nothing
				elseif vTaxaJuros = 10 then
					'Taxa Renner
					Set RSIGPMHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & date & "') AND TAXA_ID = " & vTaxaJuros)
					Set RSIGPMVencimento = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & VencimentoParcela & "') AND TAXA_ID = " & vTaxaJuros)

					if RSIGPMHoje.EOF then
						Set RSIGPMHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & date & "') AND TAXA_ID = " & vTaxaJuros & " ORDER BY COTA_Data DESC")
					end if
								
					if RSIGPMVencimento.EOF then
						Set RSIGPMVencimento = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & VencimentoParcela & "') AND TAXA_ID = " & vTaxaJuros & " ORDER BY COTA_Data DESC")
					end if
								
					'if RSIGPMHoje.EOF or RSIGPMVencimento.EOF then
					'	vValorCorrigido = Valor
					'else
						vValorCorrigido = (RSIGPMVencimento("COTA_Indice") / RSIGPMHoje("COTA_Indice")) * Valor
					'end if
								
					RSIGPMVencimento.Close
					RSIGPMHoje.Close
					Set RSIGPMVencimento = Nothing
					Set RSIGPMHoje = Nothing
				else
					'Qualquer outro índice
					Set RSIGPMHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & date & "') AND TAXA_ID = " & vTaxaJuros)
					Set RSIGPMVencimento = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & VencimentoParcela & "') AND TAXA_ID = " & vTaxaJuros)

					if RSIGPMHoje.EOF then
						Set RSIGPMHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & date & "') AND TAXA_ID = " & vTaxaJuros & " ORDER BY COTA_Data DESC")
					end if
								
					if RSIGPMVencimento.EOF then
						Set RSIGPMVencimento = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & VencimentoParcela & "') AND TAXA_ID = " & vTaxaJuros & " ORDER BY COTA_Data DESC")
					end if
								
					'if RSIGPMHoje.EOF or RSIGPMVencimento.EOF then
					'	vValorCorrigido = Valor
					'else
						vValorCorrigido = (RSIGPMHoje("COTA_Indice") / RSIGPMVencimento("COTA_Indice")) * Valor
					'end if
								
					RSIGPMVencimento.Close
					RSIGPMHoje.Close
					Set RSIGPMVencimento = Nothing
					Set RSIGPMHoje = Nothing
				end if
			elseif vPercentJuros <> "" then
				if vNumDias > 0 then
					vValorCorrigido = (((vNumDias * (vPercentJuros / 30)) / 100) * Valor) + Valor
				end if
			end if
		end if
	end if

	Conn.Close 
	Set Conn = Nothing
	
	AtualizaValorParcelaBoleto = vValorCorrigido

end function

'Retorna o nome do funcionário de executou determinada ação
function RetornaNomeFunc(FUNC_ID)
	Dim Conn
	Dim RSFunc

	if FUNC_ID <> "" then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open Application("Conexao")

		Set RSFunc = Conn.Execute("SELECT FUNC_Nome FROM Funcionarios WITH (NOLOCK) WHERE FUNC_ID = " & FUNC_ID)
		
		if RSFunc.EOF then
			RetornaNomeFunc = "Não Encontrado"
		else
			RetornaNomeFunc = RSFunc("FUNC_Nome")
		end if
			
		RSFunc.Close
		Set RSFunc = Nothing
		Conn.Close 
		Set Conn = Nothing
	else
		RetornaNomeFunc = "Não Atribuído"
	end if
end function

'Retorna o nome do funcionário de executou determinada ação
function RetornaLogonFunc(FUNC_ID)
	Dim Conn
	Dim RSFunc

	if FUNC_ID <> "" then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open Application("Conexao")

		Set RSFunc = Conn.Execute("SELECT FUNC_Logon FROM Funcionarios WITH (NOLOCK) WHERE FUNC_ID = " & FUNC_ID)
		
		if RSFunc.EOF then
			RetornaLogonFunc = ""
		else
			RetornaLogonFunc = RSFunc("FUNC_Logon")
		end if
			
		RSFunc.Close
		Set RSFunc = Nothing
		Conn.Close 
		Set Conn = Nothing
	else
		RetornaLogonFunc = ""
	end if
end function

'Retorna o nome do recuperador que está com determinado contrato
function RetornaNomeRecup(FUNC_ID)
	Dim Conn
	Dim RSFunc

	if FUNC_ID <> "" then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.CommandTimeout = 300
		Conn.Open Application("Conexao")

		Set RSFunc = Conn.Execute("SELECT FUNC_Nome FROM Funcionarios WITH (NOLOCK) WHERE FUNC_ID = " & FUNC_ID)
		
		if RSFunc.EOF then
			RetornaNomeRecup = "Não Encontrado"
		else
			RetornaNomeRecup = RSFunc("FUNC_Nome")
		end if
			
		RSFunc.Close
		Set RSFunc = Nothing
		Conn.Close 
		Set Conn = Nothing
	else
		RetornaNomeRecup = "<font color='red'>Não Distribuído</font>"
	end if
end function



'Calcula o valor da correção de um título baseado na Politica de Negociação da empresa
function ValorCorrecao(ByVal ID_Carteira, ByVal VencimentoDebito, ByVal AtualizaAte, ByVal Valor, ByVal Sinal, ByVal PermiteAtualizacao, ByVal QtdParc, ByVal DataRecebimento, ByVal TDOC_ID, ByVal DataAtualizacao, ByVal DataVencDebito)
	Dim Conn
	Dim RSIGPMHoje
	Dim RSIGPMVencimento
	Dim RSPolitica
	Dim vAtualiza, vAtualizaDebito, vAtualizaCredito, vMulta, vPercentMulta, vMultaPrincipal, vMultaPrincipalCorrigido, vJuros, vTaxaJuros, vPercentJuros
	Dim vNumDias, vValorCorrigido, vValorMulta, vValorJuros, vValorHonorarios, vHonorarios, vPercentHonorarios, vHonorariosPrincipal, vHonorariosPrincipalCorrigido
	Dim vJurosSimples, vJurosComposto, vJurosFaixa, RSTaxaJurosFaixa, rsCont
	Dim vQtdMeses, sai, vDataNova
	Dim vVencimentoTitulo, vRecebimentoContrato, vAtualizaValor, vTaxaAtualizacao, vNumDias3
	Dim vTaxaJurosComposta
	Dim vValorMultaIOF, vValorMultaObrig, vValorIOFObrig
	Dim vCONT_ID
	Dim vNumeroCarteira
	
	vQtdMeses = 0
	vValorCorrigido = Valor
	sai = false
	vValorCorrigido = Valor
	vValorJuros = 0
	vValorMulta = 0

	Set Conn = Server.CreateObject("ADODB.Connection")
	Conn.Open Application("Conexao")

	if PermiteAtualizacao then
		Set RSPolitica = Conn.Execute("SELECT * FROM Politica_de_Negociacao p WITH (NOLOCK) JOIN REL_PolNeg_Cart r ON p.PNEG_ID = r.PNEG_ID WHERE CART_ID = " & ID_Carteira)
		Set rsCont = Conn.Execute("SELECT CONT_ID, CART_Numero FROM Carteiras WITH (NOLOCK) WHERE CART_ID = " & ID_Carteira)	
		vCONT_ID = rsCont("CONT_ID")
		vNumeroCarteira = rsCont("CART_Numero")

		if not RSPolitica.eof then
			vAtualiza = RSPolitica("PNEG_Atualiza")
			vAtualizaDebito = RSPolitica("PNEG_AtualizaDebito")
			vAtualizaCredito = RSPolitica("PNEG_AtualizaCredito")
			vMulta = RSPolitica("PNEG_Multa")
			vPercentMulta = RSPolitica("PNEG_PercentMulta")
			vMultaPrincipal = RSPolitica("PNEG_MultaPrincipal")
			vMultaPrincipalCorrigido = RSPolitica("PNEG_MultaPrincipalCorrigido")
			vAtualizaValor = RSPolitica("PNEG_AtualizaValor")
			vTaxaAtualizacao = RSPolitica("TAXA_ID_Atualizacao")
			vJuros = RSPolitica("PNEG_Juros")
			vJurosFaixa = RSPolitica("PNEG_JurosPorFaixa")
			vPercentJuros = RSPolitica("PNEG_PercentualJuros")
			vJurosSimples = RSPolitica("PNEG_JurosSimples")
			vJurosComposto = RSPolitica("PNEG_JurosComposto")
			vVencimentoTitulo = RSPolitica("PNEG_AtualizaDoVencimento")
			vRecebimentoContrato = RSPolitica("PNEG_AtualizaDoRecebimento")
			
			if vCONT_ID = 101 then 'ID_Carteira = 18 or ID_Carteira = 202 then
				vAtraso = DateDiff("d", DataVencDebito, AtualizaAte)
				if vAtraso <= 150 then
					vPercentHonorarios = 10
				elseif vAtraso >= 151 and vAtraso <= 360 and QtdParc = 1 then
					vPercentHonorarios = 15
				elseif vAtraso >= 151 and vAtraso <= 360 and QtdParc > 1 then
					vPercentHonorarios = 10
				end if
			'elseif vCONT_ID = 34 or vCONT_ID = 100 then
			'	vAtraso = DateDiff("d", DataVencDebito, AtualizaAte)
			'	if vAtraso <= 60 then
			'		vPercentHonorarios = 0
			'	end if
			end if

			if (ID_Carteira = 41 or ID_Carteira = 42 or ID_Carteira = 43 or ID_Carteira = 44) and AtualizaAte > CDate("17/10/2005") then
				AtualizaAte = CDate("17/10/2005")
			end if
			if vCONT_ID = 43 and AtualizaAte > CDate("10/04/2006") then
				AtualizaAte = CDate("10/04/2006")
			end if

			if vRecebimentoContrato then
				vNumDias = DateDiff("d", DataRecebimento, AtualizaAte)
			else
				vNumDias = DateDiff("d", VencimentoDebito, AtualizaAte)
			end if
			
			Dim vNumDias2
			vNumDias2 = DateDiff("d", DataVencDebito, AtualizaAte)

			Dim rsTaxaHoje, rsTaxaVenc
			Dim vNumMeses, vDiasRestantes, w 

			Dim rsDescTipoDoc, vDescTipoDoc
			vDescTipoDoc = ""
			if rsCont("CONT_ID") = 31 then
				Set rsDescTipoDoc = BD.Execute("SELECT TDOC_Descricao FROM Tipos_de_Documento WITH (NOLOCK) WHERE TDOC_ID = " & TDOC_ID)
				if not rsDescTipoDoc.EOF then
					vDescTipoDoc = rsDescTipoDoc("TDOC_Descricao")
				end if
				rsDescTipoDoc.Close
				Set rsDescTipoDoc =  nothing
			end if
			
			if rsCont("CONT_ID") = 83 then
				vNumDias3 = vNumDias
				if vNumDias3 > 114 then
					vNumDias3 = 114
				end if
				vValorJuros = (vValorCorrigido * (1 + (12.9 / 100))^(vNumDias3 / 30)) - vValorCorrigido
				vValorCorrigido = vValorCorrigido + vValorJuros
			elseif rsCont("CONT_ID") = 10 then
				if vDescTipoDoc = "ADIANTAMENTO DEPOSITANTE" or vDescTipoDoc = "CHEQUE ESPECIAL" or vDescTipoDoc = "CHEQUE UNIVERSITARIO MB" or vDescTipoDoc = "CHEQUE EMPRESA MB" then
					if vNumDias < 180 then
						vPercentJuros = 7
					else
						vPercentJuros = 8
					end if
				else
					if vNumDias < 180 then
						vPercentJuros = 6
					else
						vPercentJuros = 7
					end if
				end if
				vValorJuros = (vValorCorrigido * (1 + (vPercentJuros / 100))^(vNumDias / 30)) - vValorCorrigido
				vValorCorrigido = vValorCorrigido + vValorJuros
			elseif rsCont("CONT_ID") = 99 then
				vValorCorrigido = Round(Valor * (1.0002125^vNumDias), 2)
				vValorMulta = Round(vValorCorrigido * (vPercentMulta / 100), 2)
				vValorJuros = Round((Valor + vValorMulta) * ((vNumDias * (vPercentJuros / 30)) / 100), 2)
				vValorCorrigido = vValorCorrigido + vValorJuros + vValorMulta
			elseif rsCont("CONT_ID") = 50 then
				if vNumDias < 180 then
					vValorMulta = vValorCorrigido * 0.02
					vNumMeses = vNumDias \ 30
					vDiasRestantes = vNumDias mod 30
					if vNumMeses > 0 then
						For w = 1 to vNumMeses
							vValorCorrigido = vValorCorrigido * 1.099 ' 9,9% ao mês
						Next
					end if
					if vDiasRestantes > 0 then
						vValorCorrigido = vValorCorrigido * (1 + (9.9 / 3000 * vDiasRestantes))
					end if
					vValorCorrigido = vValorCorrigido + vValorMulta
				else
					Set rsTaxaVenc = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & "01/" & Month(VencimentoDebito) & "/" & Year(VencimentoDebito) & "') AND TAXA_ID = 10")
				
					if rsTaxaVenc.EOF then
						Set rsTaxaVenc = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & "01/" & Month(VencimentoDebito) & "/" & Year(VencimentoDebito) & "') AND TAXA_ID = 10 ORDER BY COTA_Data DESC")
					end if
					
					if not rsTaxaVenc.EOF then
						vValorCorrigido = vValorCorrigido * (1 + rsTaxaVenc("COTA_Indice"))
					end if

					rsTaxaVenc.Close
				end if
			elseif TDOC_ID = 67 and DateDiff("d", VencimentoDebito, AtualizaAte) < 61 then
				'Mora
				Set rsTaxaHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 4")
				Set rsTaxaVenc = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 4")

				if rsTaxaHoje.EOF then
					Set rsTaxaHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 4 ORDER BY COTA_Data DESC")
				end if
				
				if rsTaxaVenc.EOF then
					Set rsTaxaVenc = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 4 ORDER BY COTA_Data DESC")
				end if
				
				if not rsTaxaHoje.EOF and not rsTaxaVenc.EOF then
					vValorCorrigido = (rsTaxaHoje("COTA_Indice") / rsTaxaVenc("COTA_Indice")) * vValorCorrigido
				end if

				rsTaxaVenc.Close
				rsTaxaHoje.Close
				Set rsTaxaVenc = Nothing
				Set rsTaxaHoje = Nothing
			elseif TDOC_ID = 67 and DateDiff("d", VencimentoDebito, AtualizaAte) > 60 and DateDiff("d", VencimentoDebito, DataAtualizacao) < 61 then
				'Mora
				Set rsTaxaHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & DateAdd("d", 60, VencimentoDebito) & "') AND TAXA_ID = 4")
				Set rsTaxaVenc = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 4")

				if rsTaxaHoje.EOF then
					Set rsTaxaHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & DateAdd("d", 60, VencimentoDebito) & "') AND TAXA_ID = 4 ORDER BY COTA_Data DESC")
				end if
				
				if rsTaxaVenc.EOF then
					Set rsTaxaVenc = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 4 ORDER BY COTA_Data DESC")
				end if
				
				if not rsTaxaHoje.EOF and not rsTaxaVenc.EOF then
					vValorCorrigido = (rsTaxaHoje("COTA_Indice") / rsTaxaVenc("COTA_Indice")) * vValorCorrigido
				end if

				rsTaxaVenc.Close
				rsTaxaHoje.Close

				'CL/LP
				Set rsTaxaHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 5")
				Set rsTaxaVenc = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & DateAdd("d", 60, VencimentoDebito) & "') AND TAXA_ID = 5")

				if rsTaxaHoje.EOF then
					Set rsTaxaHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 5 ORDER BY COTA_Data DESC")
				end if
				
				if rsTaxaVenc.EOF then
					Set rsTaxaVenc = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & DateAdd("d", 60, VencimentoDebito) & "') AND TAXA_ID = 5 ORDER BY COTA_Data DESC")
				end if
				
				if not rsTaxaHoje.EOF and not rsTaxaVenc.EOF then
					vValorCorrigido = (rsTaxaHoje("COTA_Indice") / rsTaxaVenc("COTA_Indice")) * vValorCorrigido
				end if

				rsTaxaVenc.Close
				rsTaxaHoje.Close
				Set rsTaxaVenc = Nothing
				Set rsTaxaHoje = Nothing
				Set rsTaxaVenc = Nothing
				Set rsTaxaHoje = Nothing
			elseif (TDOC_ID = 67 and DateDiff("d", VencimentoDebito, AtualizaAte) > 60 and DateDiff("d", VencimentoDebito, DataAtualizacao) > 61) or TDOC_ID = 68 or TDOC_ID = 69 then
				'CL/LP
				Set rsTaxaHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 5")
				Set rsTaxaVenc = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 5")

				if rsTaxaHoje.EOF then
					Set rsTaxaHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 5 ORDER BY COTA_Data DESC")
				end if
				
				if rsTaxaVenc.EOF then
					Set rsTaxaVenc = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 5 ORDER BY COTA_Data DESC")
				end if
				
				if not rsTaxaHoje.EOF and not rsTaxaVenc.EOF then
					vValorCorrigido = (rsTaxaHoje("COTA_Indice") / rsTaxaVenc("COTA_Indice")) * vValorCorrigido
				end if

				rsTaxaVenc.Close
				rsTaxaHoje.Close
				Set rsTaxaVenc = Nothing
				Set rsTaxaHoje = Nothing
			else
				if vAtualiza then
					if (vAtualizaDebito and Sinal = "+") or (vAtualizaCredito and Sinal = "-") then
						if vJuros then
							vValorJuros = 0
							'Cobrar juros
							if vJurosFaixa then
								'Juros por faixa de atraso
								
								Set RSTaxaJurosFaixa = Conn.Execute("SELECT FNEG_PercentJuros, FNEG_JurosSimples, FNEG_JurosComposto FROM Faixas_de_Negociacao f WITH (NOLOCK) JOIN Tabelas_de_Negociacao t WITH (NOLOCK) ON f.TNEG_ID = t.TNEG_ID JOIN REL_TabNeg_Cart r WITH (NOLOCK) ON t.TNEG_ID = r.TNEG_ID WHERE ((" & vNumDias2 & " BETWEEN TNEG_InicioFaixaAtraso AND TNEG_FinalFaixaAtraso) OR (" & Year(DataVencDebito) & " BETWEEN TNEG_InicioAnoAtraso AND TNEG_FinalAnoAtraso)) AND CART_ID = " & ID_Carteira & " AND FNEG_QtdParcelas = " & QtdParc & " AND FNEG_PercentJuros IS NOT NULL AND FNEG_Habilitado = 1 AND TNEG_Habilitada = 1 AND TNEG_VigenciaDe <= '" & AtualizaAte & "' AND TNEG_VigenciaAte >= '" & AtualizaAte & "'")
								'Set RSTaxaJurosFaixa = Conn.Execute("SELECT FNEG_PercentJuros, FNEG_JurosSimples, FNEG_JurosComposto FROM Faixas_de_Negociacao f WITH (NOLOCK) JOIN Tabelas_de_Negociacao t WITH (NOLOCK) ON f.TNEG_ID = t.TNEG_ID JOIN REL_TabNeg_Cart r WITH (NOLOCK) ON t.TNEG_ID = r.TNEG_ID WHERE " & vNumDias2 & " BETWEEN TNEG_InicioFaixaAtraso AND TNEG_FinalFaixaAtraso AND CART_ID = " & ID_Carteira & " AND FNEG_QtdParcelas = " & QtdParc & " AND FNEG_PercentJuros IS NOT NULL AND FNEG_Habilitado = 1")
								if not RSTaxaJurosFaixa.EOF then
									vJurosSimples = RSTaxaJurosFaixa("FNEG_JurosSimples")
									vJurosComposto = RSTaxaJurosFaixa("FNEG_JurosComposto")
									vPercentJuros = RSTaxaJurosFaixa("FNEG_PercentJuros")
									
									if vJurosSimples then
										if rsCont("CONT_ID") = 31 then
											vValorJuros = (((vNumDias * Round(vPercentJuros / 30, 3)) / 100) * vValorCorrigido)
										else
											vValorJuros = (((vNumDias * (vPercentJuros / 30)) / 100) * vValorCorrigido)
										end if
									elseif vJurosComposto then
										vValorJuros = (vValorCorrigido * (1 + (vPercentJuros / 100))^(vNumDias / 30)) - vValorCorrigido
									end if
								end if							
								RSTaxaJurosFaixa.Close
								Set RSTaxaJurosFaixa = Nothing
							elseif vPercentJuros <> "" then
								if vJurosSimples then
									if rsCont("CONT_ID") = 31 then
										vValorJuros = (((vNumDias * Round(vPercentJuros / 30, 3)) / 100) * vValorCorrigido)
									else
										vValorJuros = (((vNumDias * (vPercentJuros / 30)) / 100) * vValorCorrigido)
									end if
								elseif vJurosComposto then
									vValorJuros = (vValorCorrigido * (1 + (vPercentJuros / 100))^(vNumDias / 30)) - vValorCorrigido
								end if
							end if
						end if
						if vAtualizaValor and UCASE(vDescTipoDoc) <> "CREDUCSAL" then
							if vTaxaAtualizacao = 3 then
								'IGPM
								Set RSIGPMHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 3")
								Set RSIGPMVencimento = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & VencimentoDebito & "') + 2 AND TAXA_ID = 3")

								if RSIGPMHoje.EOF then
									Set RSIGPMHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 3 ORDER BY COTA_Data DESC")
								end if
								
								if RSIGPMVencimento.EOF then
									Set RSIGPMVencimento = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & VencimentoDebito & "') AND TAXA_ID = 3 ORDER BY COTA_Data DESC")
								end if
								
								if RSIGPMHoje.EOF then
									Set RSIGPMHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE TAXA_ID = 3 ORDER BY COTA_Data DESC")
								end if
								
								if RSIGPMVencimento.EOF then
									Set RSIGPMVencimento = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE TAXA_ID = 3 ORDER BY COTA_Data DESC")
								end if

								vValorCorrigido = (RSIGPMHoje("COTA_Indice") / RSIGPMVencimento("COTA_Indice")) * vValorCorrigido
								
								RSIGPMVencimento.Close
								RSIGPMHoje.Close
								Set RSIGPMVencimento = Nothing
								Set RSIGPMHoje = Nothing
							elseif vTaxaAtualizacao = 10 then
								'IGPM
								Set RSIGPMHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = " & vTaxaAtualizacao)
								Set RSIGPMVencimento = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & VencimentoDebito & "') AND TAXA_ID = " & vTaxaAtualizacao)

								if RSIGPMHoje.EOF then
									Set RSIGPMHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
								end if
								
								if RSIGPMVencimento.EOF then
									Set RSIGPMVencimento = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & VencimentoDebito & "') AND TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
								end if
								
								if RSIGPMHoje.EOF then
									Set RSIGPMHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
								end if
								
								if RSIGPMVencimento.EOF then
									Set RSIGPMVencimento = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
								end if

								vValorCorrigido = (RSIGPMVencimento("COTA_Indice") / RSIGPMHoje("COTA_Indice")) * vValorCorrigido
								
								RSIGPMVencimento.Close
								RSIGPMHoje.Close
								Set RSIGPMVencimento = Nothing
								Set RSIGPMHoje = Nothing
							elseif vTaxaAtualizacao = 15 then
								'IGPM Pró-rata
								if CDate(AtualizaAte) > CDate(VencimentoDebito) then
									if Month(VencimentoDebito) = Month(AtualizaAte) and Year(VencimentoDebito) = Year(AtualizaAte) then
										Set RSIGPMVencimento = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & "01/" & Month(VencimentoDebito) & "/" & Year(VencimentoDebito) & "') AND TAXA_ID = " & vTaxaAtualizacao)
										if not RSIGPMVencimento.EOF then
											vValorCorrigido = DateDiff("d", VencimentoDebito, AtualizaAte) * RSIGPMVencimento("COTA_Indice") / 100 / 30 * vValorCorrigido
										end if
									else
										sai = false
										vTaxaJurosComposta = 0
										Set RSIGPMVencimento = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & "01/" & Month(VencimentoDebito) & "/" & Year(VencimentoDebito) & "') AND TAXA_ID = " & vTaxaAtualizacao)
										if not RSIGPMVencimento.EOF then
											vTaxaJurosComposta = 1 + Round((DateDiff("d", VencimentoDebito, DateAdd("m", 1, ("01/" & Month(VencimentoDebito) & "/" & Year(VencimentoDebito)))) * RSIGPMVencimento("COTA_Indice") / 100 / 30), 6)
										end if
										w = VencimentoDebito
										Do While not sai
											w = DateAdd("m", 1, w)
											if DateDiff("m", w, AtualizaAte) = 0 then
												sai = true
												Set RSIGPMVencimento = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & "01/" & Month(AtualizaAte) & "/" & Year(AtualizaAte) & "') AND TAXA_ID = " & vTaxaAtualizacao)
												if not RSIGPMVencimento.EOF then
													vTaxaJurosComposta = Round(vTaxaJurosComposta * Round((1 + (DateDiff("d", "01/" & Month(AtualizaAte) & "/" & Year(AtualizaAte), AtualizaAte) * RSIGPMVencimento("COTA_Indice") / 100 / 30)), 6), 6)
												end if
											else
												Set RSIGPMVencimento = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & "01/" & Month(w) & "/" & Year(w) & "') AND TAXA_ID = " & vTaxaAtualizacao)
												if not RSIGPMVencimento.EOF then
													vTaxaJurosComposta = Round(vTaxaJurosComposta * Round((1 + RSIGPMVencimento("COTA_Indice") / 100), 6), 6)
												end if
											end if
										Loop
										vValorCorrigido = vValorCorrigido + Round((vTaxaJurosComposta - 1) * vValorCorrigido, 2)
									end if
								end if
							else
								'outros
								Set RSIGPMHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = " & vTaxaAtualizacao)
								Set RSIGPMVencimento = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & VencimentoDebito & "') AND TAXA_ID = " & vTaxaAtualizacao)

								if RSIGPMHoje.EOF then
									Set RSIGPMHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
								end if
								
								if RSIGPMVencimento.EOF then
									Set RSIGPMVencimento = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & VencimentoDebito & "') AND TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
								end if
								
								if RSIGPMHoje.EOF then
									Set RSIGPMHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
								end if
								
								if RSIGPMVencimento.EOF then
									Set RSIGPMVencimento = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
								end if

								'if RSIGPMHoje.EOF or RSIGPMVencimento.EOF then
								'	vValorCorrigido = Valor
								'else
									vValorCorrigido = (RSIGPMHoje("COTA_Indice") / RSIGPMVencimento("COTA_Indice")) * vValorCorrigido
								'end if
								
								
								RSIGPMVencimento.Close
								RSIGPMHoje.Close
								Set RSIGPMVencimento = Nothing
								Set RSIGPMHoje = Nothing
							end if
							if rsCont("CONT_ID") = 31 then
								vValorJuros = (((vNumDias * Round(vPercentJuros / 30, 3)) / 100) * vValorCorrigido)
							end if
						end if
						vValorCorrigido = vValorCorrigido + vValorJuros
						if vMulta then
							'Cobrar multa
							vValorMulta = 0
							if vCONT_ID = 13 and vNumDias > 120 then
								vPercentMulta = 0
							end if
							if vMultaPrincipal then
								'Cobrar multa sobre o principal
								'vValorMulta = ((vNumDias * (vPercentMulta / 30)) / 100) * Valor
								vValorMulta = (vPercentMulta / 100) * Valor
							elseif vMultaPrincipalCorrigido then
								'Cobrar multa sobre o proncipal corrigido (com Juros)
								'vValorMulta = ((vNumDias * (vPercentMulta / 30)) / 100) * vValorCorrigido
								vValorMulta = (vPercentMulta / 100) * vValorCorrigido
							end if
							if rsCont("CONT_ID") = 31 then
								if UCASE(vDescTipoDoc) = "CREDUCSAL" then
									vValorMulta = (vPercentMulta / 100) * Valor
								else
									vValorCorrigido = vValorCorrigido - vValorJuros
									vValorMulta = (vPercentMulta / 100) * vValorCorrigido
									vValorCorrigido = vValorCorrigido + vValorJuros
								end if
							end if
							vValorCorrigido = vValorCorrigido + vValorMulta
						end if
					end if
				end if
			end if
		end if
	end if
	
	if vCONT_ID = 1 then
		vValorMultaObrig = Valor * 0.02
		vValorIOFObrig = (Valor + vValorMultaObrig) * 0.01
		vValorMultaIOF = vValorMultaObrig + vValorIOFObrig
		if vValorMultaIOF > vValorCorrigido - Valor then
			vValorCorrigido = Valor + vValorMultaIOF
		end if
	end if

	Conn.Close 
	Set Conn = Nothing

	ValorCorrecao = vValorCorrigido - Valor
end function



'Calcula o valor dos juros de um título baseado na Politica de Negociação da empresa
function ValorJuros(ID_Carteira, VencimentoDebito, AtualizaAte, Valor, Sinal, PermiteAtualizacao, QtdParc, DataRecebimento, TDOC_ID, DataAtualizacao, DataVencDebito)
	Dim Conn
	Dim RSIGPMHoje
	Dim RSIGPMVencimento
	Dim RSPolitica
	Dim vAtualiza, vAtualizaDebito, vAtualizaCredito, vMulta, vPercentMulta, vMultaPrincipal, vMultaPrincipalCorrigido, vJuros, vTaxaJuros, vPercentJuros
	Dim vNumDias, vValorCorrigido, vValorMulta, vValorJuros, vValorHonorarios, vHonorarios, vPercentHonorarios, vHonorariosPrincipal, vHonorariosPrincipalCorrigido
	Dim vJurosSimples, vJurosComposto, vJurosFaixa, RSTaxaJurosFaixa, rsCont
	Dim vQtdMeses, sai, vDataNova
	Dim vVencimentoTitulo, vRecebimentoContrato, vAtualizaValor, vTaxaAtualizacao
	Dim vTaxaJurosComposta
	Dim w
	
	vQtdMeses = 0
	vValorCorrigido = Valor
	sai = false
	vValorCorrigido = Valor
	vValorJuros = 0
	vValorMulta = 0
	if PermiteAtualizacao then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open Application("Conexao")
	
		Set RSPolitica = Conn.Execute("SELECT * FROM Politica_de_Negociacao p WITH (NOLOCK) JOIN REL_PolNeg_Cart r ON p.PNEG_ID = r.PNEG_ID WHERE CART_ID = " & ID_Carteira)
		Set rsCont = Conn.Execute("SELECT CONT_ID FROM Carteiras WITH (NOLOCK) WHERE CART_ID = " & ID_Carteira)	
	
		if not RSPolitica.eof then
			vAtualiza = RSPolitica("PNEG_Atualiza")
			vAtualizaDebito = RSPolitica("PNEG_AtualizaDebito")
			vAtualizaCredito = RSPolitica("PNEG_AtualizaCredito")
			vMulta = RSPolitica("PNEG_Multa")
			vPercentMulta = RSPolitica("PNEG_PercentMulta")
			vMultaPrincipal = RSPolitica("PNEG_MultaPrincipal")
			vMultaPrincipalCorrigido = RSPolitica("PNEG_MultaPrincipalCorrigido")
			vAtualizaValor = RSPolitica("PNEG_AtualizaValor")
			vTaxaAtualizacao = RSPolitica("TAXA_ID_Atualizacao")
			vJuros = RSPolitica("PNEG_Juros")
			vJurosFaixa = RSPolitica("PNEG_JurosPorFaixa")
			vPercentJuros = RSPolitica("PNEG_PercentualJuros")
			vJurosSimples = RSPolitica("PNEG_JurosSimples")
			vJurosComposto = RSPolitica("PNEG_JurosComposto")
			vVencimentoTitulo = RSPolitica("PNEG_AtualizaDoVencimento")
			vRecebimentoContrato = RSPolitica("PNEG_AtualizaDoRecebimento")
			
			'if rsCont("CONT_ID") = 3 then
			'	VencimentoDebito = DataVencDebito
			'end if

			Dim rsDescTipoDoc, vDescTipoDoc
			vDescTipoDoc = ""
			if rsCont("CONT_ID") = 31 then
				Set rsDescTipoDoc = BD.Execute("SELECT TDOC_Descricao FROM Tipos_de_Documento WITH (NOLOCK) WHERE TDOC_ID = " & TDOC_ID)
				if not rsDescTipoDoc.EOF then
					vDescTipoDoc = rsDescTipoDoc("TDOC_Descricao")
				end if
				rsDescTipoDoc.Close
				Set rsDescTipoDoc =  nothing
			end if
			
			if vRecebimentoContrato then
				vNumDias = DateDiff("d", DataRecebimento, AtualizaAte)
			else
				vNumDias = DateDiff("d", VencimentoDebito, AtualizaAte)
			end if
			
			Dim vNumDias2
			vNumDias2 = DateDiff("d", DataVencDebito, AtualizaAte)

			Dim rsTaxaHoje, rsTaxaVenc
			
			if TDOC_ID = 67 and DateDiff("d", VencimentoDebito, AtualizaAte) < 61 then
				'Mora
				Set rsTaxaHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 4")
				Set rsTaxaVenc = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 4")

				if rsTaxaHoje.EOF then
					Set rsTaxaHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 4 ORDER BY COTA_Data DESC")
				end if
				
				if rsTaxaVenc.EOF then
					Set rsTaxaVenc = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 4 ORDER BY COTA_Data DESC")
				end if
				
				if not rsTaxaHoje.EOF and not rsTaxaVenc.EOF then
					vValorCorrigido = (rsTaxaHoje("COTA_Indice") / rsTaxaVenc("COTA_Indice")) * vValorCorrigido
				end if

				rsTaxaVenc.Close
				rsTaxaHoje.Close
				Set rsTaxaVenc = Nothing
				Set rsTaxaHoje = Nothing
			elseif TDOC_ID = 67 and DateDiff("d", VencimentoDebito, AtualizaAte) > 60 and DateDiff("d", VencimentoDebito, DataAtualizacao) < 61 then
				'Mora
				Set rsTaxaHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & DateAdd("d", 60, VencimentoDebito) & "') AND TAXA_ID = 4")
				Set rsTaxaVenc = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 4")

				if rsTaxaHoje.EOF then
					Set rsTaxaHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & DateAdd("d", 60, VencimentoDebito) & "') AND TAXA_ID = 4 ORDER BY COTA_Data DESC")
				end if
				
				if rsTaxaVenc.EOF then
					Set rsTaxaVenc = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 4 ORDER BY COTA_Data DESC")
				end if
				
				if not rsTaxaHoje.EOF and not rsTaxaVenc.EOF then
					vValorCorrigido = (rsTaxaHoje("COTA_Indice") / rsTaxaVenc("COTA_Indice")) * vValorCorrigido
				end if

				rsTaxaVenc.Close
				rsTaxaHoje.Close

				'CL/LP
				Set rsTaxaHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 5")
				Set rsTaxaVenc = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & DateAdd("d", 60, VencimentoDebito) & "') AND TAXA_ID = 5")

				if rsTaxaHoje.EOF then
					Set rsTaxaHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 5 ORDER BY COTA_Data DESC")
				end if
				
				if rsTaxaVenc.EOF then
					Set rsTaxaVenc = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & DateAdd("d", 60, VencimentoDebito) & "') AND TAXA_ID = 5 ORDER BY COTA_Data DESC")
				end if
				
				if not rsTaxaHoje.EOF and not rsTaxaVenc.EOF then
					vValorCorrigido = (rsTaxaHoje("COTA_Indice") / rsTaxaVenc("COTA_Indice")) * vValorCorrigido
				end if

				rsTaxaVenc.Close
				rsTaxaHoje.Close
				Set rsTaxaVenc = Nothing
				Set rsTaxaHoje = Nothing
				Set rsTaxaVenc = Nothing
				Set rsTaxaHoje = Nothing
			elseif (TDOC_ID = 67 and DateDiff("d", VencimentoDebito, AtualizaAte) > 60 and DateDiff("d", VencimentoDebito, DataAtualizacao) > 61) or TDOC_ID = 68 or TDOC_ID = 69 then
				'CL/LP
				Set rsTaxaHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 5")
				Set rsTaxaVenc = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 5")

				if rsTaxaHoje.EOF then
					Set rsTaxaHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 5 ORDER BY COTA_Data DESC")
				end if
				
				if rsTaxaVenc.EOF then
					Set rsTaxaVenc = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 5 ORDER BY COTA_Data DESC")
				end if
				
				if not rsTaxaHoje.EOF and not rsTaxaVenc.EOF then
					vValorCorrigido = (rsTaxaHoje("COTA_Indice") / rsTaxaVenc("COTA_Indice")) * vValorCorrigido
				end if

				rsTaxaVenc.Close
				rsTaxaHoje.Close
				Set rsTaxaVenc = Nothing
				Set rsTaxaHoje = Nothing
			else
				if vAtualiza then
					if (vAtualizaDebito and Sinal = "+") or (vAtualizaCredito and Sinal = "-") then
						if vJuros then
							vValorJuros = 0
							'Cobrar juros
							if vJurosFaixa then
								'Juros por faixa de atraso
								'Set RSTaxaJurosFaixa = Conn.Execute("SELECT FNEG_PercentJuros, FNEG_JurosSimples, FNEG_JurosComposto FROM Faixas_de_Negociacao f WITH (NOLOCK) JOIN Tabelas_de_Negociacao t WITH (NOLOCK) ON f.TNEG_ID = t.TNEG_ID JOIN REL_TabNeg_Cart r WITH (NOLOCK) ON t.TNEG_ID = r.TNEG_ID WHERE " & vNumDias2 & " BETWEEN TNEG_InicioFaixaAtraso AND TNEG_FinalFaixaAtraso AND CART_ID = " & ID_Carteira & " AND FNEG_QtdParcelas = " & QtdParc & " AND FNEG_PercentJuros IS NOT NULL AND FNEG_Habilitado = 1")
								Set RSTaxaJurosFaixa = Conn.Execute("SELECT FNEG_PercentJuros, FNEG_JurosSimples, FNEG_JurosComposto FROM Faixas_de_Negociacao f WITH (NOLOCK) JOIN Tabelas_de_Negociacao t WITH (NOLOCK) ON f.TNEG_ID = t.TNEG_ID JOIN REL_TabNeg_Cart r WITH (NOLOCK) ON t.TNEG_ID = r.TNEG_ID WHERE ((" & vNumDias2 & " BETWEEN TNEG_InicioFaixaAtraso AND TNEG_FinalFaixaAtraso) OR (" & Year(DataVencDebito) & " BETWEEN TNEG_InicioAnoAtraso AND TNEG_FinalAnoAtraso)) AND CART_ID = " & ID_Carteira & " AND FNEG_QtdParcelas = " & QtdParc & " AND FNEG_PercentJuros IS NOT NULL AND FNEG_Habilitado = 1 AND TNEG_Habilitada = 1 AND TNEG_VigenciaDe <= '" & AtualizaAte & "' AND TNEG_VigenciaAte >= '" & AtualizaAte & "'")
								if not RSTaxaJurosFaixa.EOF then
									vJurosSimples = RSTaxaJurosFaixa("FNEG_JurosSimples")
									vJurosComposto = RSTaxaJurosFaixa("FNEG_JurosComposto")
									vPercentJuros = RSTaxaJurosFaixa("FNEG_PercentJuros")
									
									if vJurosSimples then
										if rsCont("CONT_ID") = 31 then
											vValorJuros = (((vNumDias * Round(vPercentJuros / 30, 3)) / 100) * vValorCorrigido)
										else
											vValorJuros = (((vNumDias * (vPercentJuros / 30)) / 100) * vValorCorrigido)
										end if
									elseif vJurosComposto then
										vValorJuros = (vValorCorrigido * (1 + (vPercentJuros / 100))^(vNumDias / 30)) - vValorCorrigido
									end if
								end if							
								RSTaxaJurosFaixa.Close
								Set RSTaxaJurosFaixa = Nothing
							elseif vPercentJuros <> "" then
								if vJurosSimples then
									if rsCont("CONT_ID") = 31 then
										vValorJuros = (((vNumDias * Round(vPercentJuros / 30, 3)) / 100) * vValorCorrigido)
									else
										vValorJuros = (((vNumDias * (vPercentJuros / 30)) / 100) * vValorCorrigido)
									end if
								elseif vJurosComposto then
									vValorJuros = (vValorCorrigido * (1 + (vPercentJuros / 100))^(vNumDias / 30)) - vValorCorrigido
								end if
							end if
						end if
						if vAtualizaValor and UCASE(vDescTipoDoc) <> "CREDUCSAL" then
							if vTaxaAtualizacao = 3 then
								'IGPM
								Set RSIGPMHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 3")
								Set RSIGPMVencimento = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & VencimentoDebito & "') + 2 AND TAXA_ID = 3")

								if RSIGPMHoje.EOF then
									Set RSIGPMHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 3 ORDER BY COTA_Data DESC")
								end if
								
								if RSIGPMVencimento.EOF then
									Set RSIGPMVencimento = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & VencimentoDebito & "') AND TAXA_ID = 3 ORDER BY COTA_Data DESC")
								end if
								
								'if RSIGPMHoje.EOF or RSIGPMVencimento.EOF then
								'	vValorCorrigido = Valor
								'else
									vValorCorrigido = (RSIGPMHoje("COTA_Indice") / RSIGPMVencimento("COTA_Indice")) * vValorCorrigido
								'end if
								
								
								RSIGPMVencimento.Close
								RSIGPMHoje.Close
								Set RSIGPMVencimento = Nothing
								Set RSIGPMHoje = Nothing
							elseif vTaxaAtualizacao = 10 then
								'Taxa Renner
								Set RSIGPMHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = " & vTaxaAtualizacao)
								Set RSIGPMVencimento = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & VencimentoDebito & "') AND TAXA_ID = " & vTaxaAtualizacao)

								if RSIGPMHoje.EOF then
									Set RSIGPMHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
								end if
								
								if RSIGPMVencimento.EOF then
									Set RSIGPMVencimento = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & VencimentoDebito & "') AND TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
								end if
								
								if RSIGPMHoje.EOF then
									Set RSIGPMHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
								end if
								
								if RSIGPMVencimento.EOF then
									Set RSIGPMVencimento = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
								end if

								vValorCorrigido = (RSIGPMVencimento("COTA_Indice") / RSIGPMHoje("COTA_Indice")) * vValorCorrigido
								
								RSIGPMVencimento.Close
								RSIGPMHoje.Close
								Set RSIGPMVencimento = Nothing
								Set RSIGPMHoje = Nothing
							elseif vTaxaAtualizacao = 15 then
								'IGPM Pró-rata
								if CDate(AtualizaAte) > CDate(VencimentoDebito) then
									if Month(VencimentoDebito) = Month(AtualizaAte) and Year(VencimentoDebito) = Year(AtualizaAte) then
										Set RSIGPMVencimento = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & "01/" & Month(VencimentoDebito) & "/" & Year(VencimentoDebito) & "') AND TAXA_ID = " & vTaxaAtualizacao)
										if not RSIGPMVencimento.EOF then
											vValorCorrigido = DateDiff("d", VencimentoDebito, AtualizaAte) * RSIGPMVencimento("COTA_Indice") / 100 / 30 * vValorCorrigido
										end if
									else
										sai = false
										vTaxaJurosComposta = 0
										Set RSIGPMVencimento = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & "01/" & Month(VencimentoDebito) & "/" & Year(VencimentoDebito) & "') AND TAXA_ID = " & vTaxaAtualizacao)
										if not RSIGPMVencimento.EOF then
											vTaxaJurosComposta = 1 + Round((DateDiff("d", VencimentoDebito, DateAdd("m", 1, ("01/" & Month(VencimentoDebito) & "/" & Year(VencimentoDebito)))) * RSIGPMVencimento("COTA_Indice") / 100 / 30), 6)
										end if
										w = VencimentoDebito
										Do While not sai
											w = DateAdd("m", 1, w)
											if DateDiff("m", w, AtualizaAte) = 0 then
												sai = true
												Set RSIGPMVencimento = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & "01/" & Month(AtualizaAte) & "/" & Year(AtualizaAte) & "') AND TAXA_ID = " & vTaxaAtualizacao)
												if not RSIGPMVencimento.EOF then
													vTaxaJurosComposta = Round(vTaxaJurosComposta * Round((1 + (DateDiff("d", "01/" & Month(AtualizaAte) & "/" & Year(AtualizaAte), AtualizaAte) * RSIGPMVencimento("COTA_Indice") / 100 / 30)), 6), 6)
												end if
											else
												Set RSIGPMVencimento = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & "01/" & Month(w) & "/" & Year(w) & "') AND TAXA_ID = " & vTaxaAtualizacao)
												if not RSIGPMVencimento.EOF then
													vTaxaJurosComposta = Round(vTaxaJurosComposta * Round((1 + RSIGPMVencimento("COTA_Indice") / 100), 6), 6)
												end if
											end if
										Loop
										vValorCorrigido = vValorCorrigido + Round((vTaxaJurosComposta - 1) * vValorCorrigido, 2)
									end if
								end if
							else
								'outros
								Set RSIGPMHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = " & vTaxaAtualizacao)
								Set RSIGPMVencimento = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & VencimentoDebito & "') AND TAXA_ID = " & vTaxaAtualizacao)

								if RSIGPMHoje.EOF then
									Set RSIGPMHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
								end if
								
								if RSIGPMVencimento.EOF then
									Set RSIGPMVencimento = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & VencimentoDebito & "') AND TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
								end if
								
								if RSIGPMHoje.EOF then
									Set RSIGPMHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
								end if
								
								if RSIGPMVencimento.EOF then
									Set RSIGPMVencimento = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
								end if

								'if RSIGPMHoje.EOF or RSIGPMVencimento.EOF then
								'	vValorCorrigido = Valor
								'else
									vValorCorrigido = (RSIGPMHoje("COTA_Indice") / RSIGPMVencimento("COTA_Indice")) * vValorCorrigido
								'end if
								
								
								RSIGPMVencimento.Close
								RSIGPMHoje.Close
								Set RSIGPMVencimento = Nothing
								Set RSIGPMHoje = Nothing
							end if
							if rsCont("CONT_ID") = 31 then
								vValorJuros = (((vNumDias * Round(vPercentJuros / 30, 3)) / 100) * vValorCorrigido)
							end if
						end if
						vValorCorrigido = vValorCorrigido + vValorJuros
					end if
				end if
			end if
		end if

		Conn.Close 
		Set Conn = Nothing
	end if
	
	ValorJuros = vValorCorrigido - Valor
end function

'Calcula o valor dos juros de um título baseado na Politica de Negociação da empresa
function ValorJurosManual(ID_Carteira, VencimentoDebito, AtualizaAte, Valor, Sinal, PermiteAtualizacao, QtdParc, DataRecebimento, Juros, TDOC_ID, DataAtualizacao)
	Dim Conn
	Dim RSIGPMHoje
	Dim RSIGPMVencimento
	Dim RSPolitica
	Dim vAtualiza, vAtualizaDebito, vAtualizaCredito, vMulta, vPercentMulta, vMultaPrincipal, vMultaPrincipalCorrigido, vJuros, vTaxaJuros, vPercentJuros
	Dim vNumDias, vValorCorrigido, vValorMulta, vValorJuros, vValorHonorarios, vHonorarios, vPercentHonorarios, vHonorariosPrincipal, vHonorariosPrincipalCorrigido
	Dim vJurosSimples, vJurosComposto, vJurosFaixa, RSTaxaJurosFaixa, rsCont
	Dim vQtdMeses, sai, vDataNova
	Dim vVencimentoTitulo, vRecebimentoContrato, vAtualizaValor, vTaxaAtualizacao
	
	vQtdMeses = 0
	vValorCorrigido = Valor
	sai = false
	vValorCorrigido = Valor
	vValorJuros = 0
	vValorMulta = 0
	if PermiteAtualizacao then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open Application("Conexao")
	
		Set RSPolitica = Conn.Execute("SELECT * FROM Politica_de_Negociacao p WITH (NOLOCK) JOIN REL_PolNeg_Cart r ON p.PNEG_ID = r.PNEG_ID WHERE CART_ID = " & ID_Carteira)
		Set rsCont = Conn.Execute("SELECT CONT_ID FROM Carteiras WITH (NOLOCK) WHERE CART_ID = " & ID_Carteira)	
	
		if not RSPolitica.eof then
			vAtualiza = RSPolitica("PNEG_Atualiza")
			vAtualizaDebito = RSPolitica("PNEG_AtualizaDebito")
			vAtualizaCredito = RSPolitica("PNEG_AtualizaCredito")
			vMulta = RSPolitica("PNEG_Multa")
			vPercentMulta = RSPolitica("PNEG_PercentMulta")
			vMultaPrincipal = RSPolitica("PNEG_MultaPrincipal")
			vMultaPrincipalCorrigido = RSPolitica("PNEG_MultaPrincipalCorrigido")
			vAtualizaValor = RSPolitica("PNEG_AtualizaValor")
			vTaxaAtualizacao = RSPolitica("TAXA_ID_Atualizacao")
			vJuros = RSPolitica("PNEG_Juros")
			vJurosFaixa = RSPolitica("PNEG_JurosPorFaixa")
			vPercentJuros = Juros
			vJurosSimples = RSPolitica("PNEG_JurosSimples")
			vJurosComposto = RSPolitica("PNEG_JurosComposto")
			vVencimentoTitulo = RSPolitica("PNEG_AtualizaDoVencimento")
			vRecebimentoContrato = RSPolitica("PNEG_AtualizaDoRecebimento")
			
			'if rsCont("CONT_ID") = 3 then
			'	VencimentoDebito = DataVencDebito
			'end if

			if vRecebimentoContrato then
				vNumDias = DateDiff("d", DataRecebimento, AtualizaAte)
			else
				vNumDias = DateDiff("d", VencimentoDebito, AtualizaAte)
			end if
			
			Dim rsTaxaHoje, rsTaxaVenc
			
			if TDOC_ID = 67 and DateDiff("d", VencimentoDebito, AtualizaAte) < 61 then
				'Mora
				Set rsTaxaHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 4")
				Set rsTaxaVenc = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 4")

				if rsTaxaHoje.EOF then
					Set rsTaxaHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 4 ORDER BY COTA_Data DESC")
				end if
				
				if rsTaxaVenc.EOF then
					Set rsTaxaVenc = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 4 ORDER BY COTA_Data DESC")
				end if
				
				if not rsTaxaHoje.EOF and not rsTaxaVenc.EOF then
					vValorCorrigido = (rsTaxaHoje("COTA_Indice") / rsTaxaVenc("COTA_Indice")) * vValorCorrigido
				end if

				rsTaxaVenc.Close
				rsTaxaHoje.Close
				Set rsTaxaVenc = Nothing
				Set rsTaxaHoje = Nothing
			elseif TDOC_ID = 67 and DateDiff("d", VencimentoDebito, AtualizaAte) > 60 and DateDiff("d", VencimentoDebito, DataAtualizacao) < 61 then
				'Mora
				Set rsTaxaHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & DateAdd("d", 60, VencimentoDebito) & "') AND TAXA_ID = 4")
				Set rsTaxaVenc = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 4")

				if rsTaxaHoje.EOF then
					Set rsTaxaHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & DateAdd("d", 60, VencimentoDebito) & "') AND TAXA_ID = 4 ORDER BY COTA_Data DESC")
				end if
				
				if rsTaxaVenc.EOF then
					Set rsTaxaVenc = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 4 ORDER BY COTA_Data DESC")
				end if
				
				if not rsTaxaHoje.EOF and not rsTaxaVenc.EOF then
					vValorCorrigido = (rsTaxaHoje("COTA_Indice") / rsTaxaVenc("COTA_Indice")) * vValorCorrigido
				end if

				rsTaxaVenc.Close
				rsTaxaHoje.Close

				'CL/LP
				Set rsTaxaHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 5")
				Set rsTaxaVenc = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & DateAdd("d", 60, VencimentoDebito) & "') AND TAXA_ID = 5")

				if rsTaxaHoje.EOF then
					Set rsTaxaHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 5 ORDER BY COTA_Data DESC")
				end if
				
				if rsTaxaVenc.EOF then
					Set rsTaxaVenc = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & DateAdd("d", 60, VencimentoDebito) & "') AND TAXA_ID = 5 ORDER BY COTA_Data DESC")
				end if
				
				if not rsTaxaHoje.EOF and not rsTaxaVenc.EOF then
					vValorCorrigido = (rsTaxaHoje("COTA_Indice") / rsTaxaVenc("COTA_Indice")) * vValorCorrigido
				end if

				rsTaxaVenc.Close
				rsTaxaHoje.Close
				Set rsTaxaVenc = Nothing
				Set rsTaxaHoje = Nothing
				Set rsTaxaVenc = Nothing
				Set rsTaxaHoje = Nothing
			elseif (TDOC_ID = 67 and DateDiff("d", VencimentoDebito, AtualizaAte) > 60 and DateDiff("d", VencimentoDebito, DataAtualizacao) > 61) or TDOC_ID = 68 or TDOC_ID = 69 then
				'CL/LP
				Set rsTaxaHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 5")
				Set rsTaxaVenc = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 5")

				if rsTaxaHoje.EOF then
					Set rsTaxaHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 5 ORDER BY COTA_Data DESC")
				end if
				
				if rsTaxaVenc.EOF then
					Set rsTaxaVenc = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 5 ORDER BY COTA_Data DESC")
				end if
				
				if not rsTaxaHoje.EOF and not rsTaxaVenc.EOF then
					vValorCorrigido = (rsTaxaHoje("COTA_Indice") / rsTaxaVenc("COTA_Indice")) * vValorCorrigido
				end if

				rsTaxaVenc.Close
				rsTaxaHoje.Close
				Set rsTaxaVenc = Nothing
				Set rsTaxaHoje = Nothing
			else
				if vAtualiza then
					if (vAtualizaDebito and Sinal = "+") or (vAtualizaCredito and Sinal = "-") then
						if vJuros then
							vValorJuros = 0
							'Cobrar juros
							if vPercentJuros <> "" then
								if vJurosSimples then
									vValorJuros = (((vNumDias * (vPercentJuros / 30)) / 100) * vValorCorrigido)
								elseif vJurosComposto then
									vValorJuros = vValorCorrigido * ((vPercentJuros / 100))^(vNumDias / 30) 
								end if
							end if
						end if
						if vAtualizaValor then
							if vTaxaAtualizacao = 3 then
								'IGPM
								Set RSIGPMHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 3")
								Set RSIGPMVencimento = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & VencimentoDebito & "') + 2 AND TAXA_ID = 3")

								if RSIGPMHoje.EOF then
									Set RSIGPMHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 3 ORDER BY COTA_Data DESC")
								end if
								
								if RSIGPMVencimento.EOF then
									Set RSIGPMVencimento = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & VencimentoDebito & "') AND TAXA_ID = 3 ORDER BY COTA_Data DESC")
								end if
								
								'if RSIGPMHoje.EOF or RSIGPMVencimento.EOF then
								'	vValorCorrigido = Valor
								'else
									vValorCorrigido = (RSIGPMHoje("COTA_Indice") / RSIGPMVencimento("COTA_Indice")) * vValorCorrigido
								'end if
								
								
								RSIGPMVencimento.Close
								RSIGPMHoje.Close
								Set RSIGPMVencimento = Nothing
								Set RSIGPMHoje = Nothing
							elseif vTaxaAtualizacao = 10 then
								'Taxa Renner
								Set RSIGPMHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = " & vTaxaAtualizacao)
								Set RSIGPMVencimento = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & VencimentoDebito & "') AND TAXA_ID = " & vTaxaAtualizacao)

								if RSIGPMHoje.EOF then
									Set RSIGPMHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
								end if
								
								if RSIGPMVencimento.EOF then
									Set RSIGPMVencimento = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & VencimentoDebito & "') AND TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
								end if
								
								if RSIGPMHoje.EOF then
									Set RSIGPMHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
								end if
								
								if RSIGPMVencimento.EOF then
									Set RSIGPMVencimento = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
								end if

								vValorCorrigido = (RSIGPMVencimento("COTA_Indice") / RSIGPMHoje("COTA_Indice")) * vValorCorrigido
								
								RSIGPMVencimento.Close
								RSIGPMHoje.Close
								Set RSIGPMVencimento = Nothing
								Set RSIGPMHoje = Nothing
							else
								'outros
								Set RSIGPMHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = " & vTaxaAtualizacao)
								Set RSIGPMVencimento = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & VencimentoDebito & "') AND TAXA_ID = " & vTaxaAtualizacao)

								if RSIGPMHoje.EOF then
									Set RSIGPMHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
								end if
								
								if RSIGPMVencimento.EOF then
									Set RSIGPMVencimento = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & VencimentoDebito & "') AND TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
								end if
								
								if RSIGPMHoje.EOF then
									Set RSIGPMHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
								end if
								
								if RSIGPMVencimento.EOF then
									Set RSIGPMVencimento = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
								end if

								'if RSIGPMHoje.EOF or RSIGPMVencimento.EOF then
								'	vValorCorrigido = Valor
								'else
									vValorCorrigido = (RSIGPMHoje("COTA_Indice") / RSIGPMVencimento("COTA_Indice")) * vValorCorrigido
								'end if
								
								
								RSIGPMVencimento.Close
								RSIGPMHoje.Close
								Set RSIGPMVencimento = Nothing
								Set RSIGPMHoje = Nothing
							end if
						end if
						vValorCorrigido = vValorCorrigido + vValorJuros
					end if
				end if
			end if
		end if

		Conn.Close 
		Set Conn = Nothing
	end if
	
	ValorJurosManual = vValorCorrigido - Valor
end function

'Calcula o valor da correção de um título baseado na Politica de Negociação da empresa
function ValorMulta(ID_Carteira, VencimentoDebito, AtualizaAte, Valor, Sinal, PermiteAtualizacao, QtdParc, DataRecebimento, TDOC_ID, DataAtualizacao, DataVencDebito)
	Dim Conn
	Dim RSIGPMHoje
	Dim RSIGPMVencimento
	Dim RSPolitica
	Dim vAtualiza, vAtualizaDebito, vAtualizaCredito, vMulta, vPercentMulta, vMultaPrincipal, vMultaPrincipalCorrigido, vJuros, vTaxaJuros, vPercentJuros
	Dim vNumDias, vValorCorrigido, vValorMulta, vValorJuros, vValorHonorarios, vHonorarios, vPercentHonorarios, vHonorariosPrincipal, vHonorariosPrincipalCorrigido
	Dim vJurosSimples, vJurosComposto, vJurosFaixa, RSTaxaJurosFaixa, rsCont
	Dim vQtdMeses, sai, vDataNova
	Dim vVencimentoTitulo, vRecebimentoContrato, vAtualizaValor, vTaxaAtualizacao
	Dim vTaxaJurosComposta, w
	
	vQtdMeses = 0
	vValorCorrigido = Valor
	sai = false
	vValorCorrigido = Valor
	vValorJuros = 0
	vValorMulta = 0
	if PermiteAtualizacao then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open Application("Conexao")
	
		Set RSPolitica = Conn.Execute("SELECT * FROM Politica_de_Negociacao p WITH (NOLOCK) JOIN REL_PolNeg_Cart r ON p.PNEG_ID = r.PNEG_ID WHERE CART_ID = " & ID_Carteira)
		Set rsCont = Conn.Execute("SELECT CONT_ID FROM Carteiras WITH (NOLOCK) WHERE CART_ID = " & ID_Carteira)	
	
		if not RSPolitica.eof then
			vAtualiza = RSPolitica("PNEG_Atualiza")
			vAtualizaDebito = RSPolitica("PNEG_AtualizaDebito")
			vAtualizaCredito = RSPolitica("PNEG_AtualizaCredito")
			vMulta = RSPolitica("PNEG_Multa")
			vPercentMulta = RSPolitica("PNEG_PercentMulta")
			vMultaPrincipal = RSPolitica("PNEG_MultaPrincipal")
			vMultaPrincipalCorrigido = RSPolitica("PNEG_MultaPrincipalCorrigido")
			vAtualizaValor = RSPolitica("PNEG_AtualizaValor")
			vTaxaAtualizacao = RSPolitica("TAXA_ID_Atualizacao")
			vJuros = RSPolitica("PNEG_Juros")
			vJurosFaixa = RSPolitica("PNEG_JurosPorFaixa")
			vPercentJuros = RSPolitica("PNEG_PercentualJuros")
			vJurosSimples = RSPolitica("PNEG_JurosSimples")
			vJurosComposto = RSPolitica("PNEG_JurosComposto")
			vVencimentoTitulo = RSPolitica("PNEG_AtualizaDoVencimento")
			vRecebimentoContrato = RSPolitica("PNEG_AtualizaDoRecebimento")
			
			'if rsCont("CONT_ID") = 3 then
			'	VencimentoDebito = DataVencDebito
			'end if

			Dim rsDescTipoDoc, vDescTipoDoc
			vDescTipoDoc = ""
			if rsCont("CONT_ID") = 31 then
				Set rsDescTipoDoc = BD.Execute("SELECT TDOC_Descricao FROM Tipos_de_Documento WITH (NOLOCK) WHERE TDOC_ID = " & TDOC_ID)
				if not rsDescTipoDoc.EOF then
					vDescTipoDoc = rsDescTipoDoc("TDOC_Descricao")
				end if
				rsDescTipoDoc.Close
				Set rsDescTipoDoc =  nothing
			end if
			
			if vRecebimentoContrato then
				vNumDias = DateDiff("d", DataRecebimento, AtualizaAte)
			else
				vNumDias = DateDiff("d", VencimentoDebito, AtualizaAte)
			end if
			
			Dim vNumDias2
			vNumDias2 = DateDiff("d", DataVencDebito, AtualizaAte)

			Dim rsTaxaHoje, rsTaxaVenc
			
			if TDOC_ID = 67 and DateDiff("d", VencimentoDebito, AtualizaAte) < 61 then
				'Mora
				Set rsTaxaHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 4")
				Set rsTaxaVenc = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 4")

				if rsTaxaHoje.EOF then
					Set rsTaxaHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 4 ORDER BY COTA_Data DESC")
				end if
				
				if rsTaxaVenc.EOF then
					Set rsTaxaVenc = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 4 ORDER BY COTA_Data DESC")
				end if
				
				if not rsTaxaHoje.EOF and not rsTaxaVenc.EOF then
					vValorCorrigido = (rsTaxaHoje("COTA_Indice") / rsTaxaVenc("COTA_Indice")) * vValorCorrigido
				end if

				rsTaxaVenc.Close
				rsTaxaHoje.Close
				Set rsTaxaVenc = Nothing
				Set rsTaxaHoje = Nothing
			elseif TDOC_ID = 67 and DateDiff("d", VencimentoDebito, AtualizaAte) > 60 and DateDiff("d", VencimentoDebito, DataAtualizacao) < 61 then
				'Mora
				Set rsTaxaHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & DateAdd("d", 60, VencimentoDebito) & "') AND TAXA_ID = 4")
				Set rsTaxaVenc = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 4")

				if rsTaxaHoje.EOF then
					Set rsTaxaHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & DateAdd("d", 60, VencimentoDebito) & "') AND TAXA_ID = 4 ORDER BY COTA_Data DESC")
				end if
				
				if rsTaxaVenc.EOF then
					Set rsTaxaVenc = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 4 ORDER BY COTA_Data DESC")
				end if
				
				if not rsTaxaHoje.EOF and not rsTaxaVenc.EOF then
					vValorCorrigido = (rsTaxaHoje("COTA_Indice") / rsTaxaVenc("COTA_Indice")) * vValorCorrigido
				end if

				rsTaxaVenc.Close
				rsTaxaHoje.Close

				'CL/LP
				Set rsTaxaHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 5")
				Set rsTaxaVenc = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & DateAdd("d", 60, VencimentoDebito) & "') AND TAXA_ID = 5")

				if rsTaxaHoje.EOF then
					Set rsTaxaHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 5 ORDER BY COTA_Data DESC")
				end if
				
				if rsTaxaVenc.EOF then
					Set rsTaxaVenc = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & DateAdd("d", 60, VencimentoDebito) & "') AND TAXA_ID = 5 ORDER BY COTA_Data DESC")
				end if
				
				if not rsTaxaHoje.EOF and not rsTaxaVenc.EOF then
					vValorCorrigido = (rsTaxaHoje("COTA_Indice") / rsTaxaVenc("COTA_Indice")) * vValorCorrigido
				end if

				rsTaxaVenc.Close
				rsTaxaHoje.Close
				Set rsTaxaVenc = Nothing
				Set rsTaxaHoje = Nothing
				Set rsTaxaVenc = Nothing
				Set rsTaxaHoje = Nothing
			elseif (TDOC_ID = 67 and DateDiff("d", VencimentoDebito, AtualizaAte) > 60 and DateDiff("d", VencimentoDebito, DataAtualizacao) > 61) or TDOC_ID = 68 or TDOC_ID = 69 then
				'CL/LP
				Set rsTaxaHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 5")
				Set rsTaxaVenc = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 5")

				if rsTaxaHoje.EOF then
					Set rsTaxaHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 5 ORDER BY COTA_Data DESC")
				end if
				
				if rsTaxaVenc.EOF then
					Set rsTaxaVenc = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 5 ORDER BY COTA_Data DESC")
				end if
				
				if not rsTaxaHoje.EOF and not rsTaxaVenc.EOF then
					vValorCorrigido = (rsTaxaHoje("COTA_Indice") / rsTaxaVenc("COTA_Indice")) * vValorCorrigido
				end if

				rsTaxaVenc.Close
				rsTaxaHoje.Close
				Set rsTaxaVenc = Nothing
				Set rsTaxaHoje = Nothing
			else
				if vAtualiza then
					if (vAtualizaDebito and Sinal = "+") or (vAtualizaCredito and Sinal = "-") then
						if vJuros then
							vValorJuros = 0
							'Cobrar juros
							if vJurosFaixa then
								'Juros por faixa de atraso
								Set RSTaxaJurosFaixa = Conn.Execute("SELECT FNEG_PercentJuros, FNEG_JurosSimples, FNEG_JurosComposto FROM Faixas_de_Negociacao f WITH (NOLOCK) JOIN Tabelas_de_Negociacao t WITH (NOLOCK) ON f.TNEG_ID = t.TNEG_ID JOIN REL_TabNeg_Cart r WITH (NOLOCK) ON t.TNEG_ID = r.TNEG_ID WHERE ((" & vNumDias2 & " BETWEEN TNEG_InicioFaixaAtraso AND TNEG_FinalFaixaAtraso) OR (" & Year(DataVencDebito) & " BETWEEN TNEG_InicioAnoAtraso AND TNEG_FinalAnoAtraso)) AND CART_ID = " & ID_Carteira & " AND FNEG_QtdParcelas = " & QtdParc & " AND FNEG_PercentJuros IS NOT NULL AND FNEG_Habilitado = 1 AND TNEG_Habilitada = 1 AND TNEG_VigenciaDe <= '" & AtualizaAte & "' AND TNEG_VigenciaAte >= '" & AtualizaAte & "'")
								'Set RSTaxaJurosFaixa = Conn.Execute("SELECT FNEG_PercentJuros, FNEG_JurosSimples, FNEG_JurosComposto FROM Faixas_de_Negociacao f WITH (NOLOCK) JOIN Tabelas_de_Negociacao t WITH (NOLOCK) ON f.TNEG_ID = t.TNEG_ID JOIN REL_TabNeg_Cart r WITH (NOLOCK) ON t.TNEG_ID = r.TNEG_ID WHERE " & vNumDias2 & " BETWEEN TNEG_InicioFaixaAtraso AND TNEG_FinalFaixaAtraso AND CART_ID = " & ID_Carteira & " AND FNEG_QtdParcelas = " & QtdParc & " AND FNEG_PercentJuros IS NOT NULL AND FNEG_Habilitado = 1")
								if not RSTaxaJurosFaixa.EOF then
									vJurosSimples = RSTaxaJurosFaixa("FNEG_JurosSimples")
									vJurosComposto = RSTaxaJurosFaixa("FNEG_JurosComposto")
									vPercentJuros = RSTaxaJurosFaixa("FNEG_PercentJuros")
									
									if vJurosSimples then
										if rsCont("CONT_ID") = 31 then
											vValorJuros = (((vNumDias * Round(vPercentJuros / 30, 3)) / 100) * vValorCorrigido)
										else
											vValorJuros = (((vNumDias * (vPercentJuros / 30)) / 100) * vValorCorrigido)
										end if
									elseif vJurosComposto then
										vValorJuros = (vValorCorrigido * (1 + (vPercentJuros / 100))^(vNumDias / 30)) - vValorCorrigido
									end if
								end if
								RSTaxaJurosFaixa.Close
								Set RSTaxaJurosFaixa = Nothing
							elseif vPercentJuros <> "" then
								if vJurosSimples then
									if rsCont("CONT_ID") = 31 then
										vValorJuros = (((vNumDias * Round(vPercentJuros / 30, 3)) / 100) * vValorCorrigido)
									else
										vValorJuros = (((vNumDias * (vPercentJuros / 30)) / 100) * vValorCorrigido)
									end if
								elseif vJurosComposto then
									vValorJuros = (vValorCorrigido * (1 + (vPercentJuros / 100))^(vNumDias / 30)) - vValorCorrigido
								end if
							end if
						end if
						if vAtualizaValor and UCASE(vDescTipoDoc) <> "CREDUCSAL" then
							if vTaxaAtualizacao = 3 then
								'IGPM
								Set RSIGPMHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 3")
								Set RSIGPMVencimento = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & VencimentoDebito & "') + 2 AND TAXA_ID = 3")

								if RSIGPMHoje.EOF then
									Set RSIGPMHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 3 ORDER BY COTA_Data DESC")
								end if
								
								if RSIGPMVencimento.EOF then
									Set RSIGPMVencimento = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & VencimentoDebito & "') AND TAXA_ID = 3 ORDER BY COTA_Data DESC")
								end if
								
								if RSIGPMHoje.EOF then
									Set RSIGPMHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE TAXA_ID = 3 ORDER BY COTA_Data DESC")
								end if
								
								if RSIGPMVencimento.EOF then
									Set RSIGPMVencimento = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE TAXA_ID = 3 ORDER BY COTA_Data DESC")
								end if

								'if RSIGPMHoje.EOF or RSIGPMVencimento.EOF then
								'	vValorCorrigido = Valor
								'else
									vValorCorrigido = (RSIGPMHoje("COTA_Indice") / RSIGPMVencimento("COTA_Indice")) * vValorCorrigido
								'end if
								
								
								RSIGPMVencimento.Close
								RSIGPMHoje.Close
								Set RSIGPMVencimento = Nothing
								Set RSIGPMHoje = Nothing
							elseif vTaxaAtualizacao = 10 then
								'IGPM
								Set RSIGPMHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = " & vTaxaAtualizacao)
								Set RSIGPMVencimento = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & VencimentoDebito & "') AND TAXA_ID = " & vTaxaAtualizacao)

								if RSIGPMHoje.EOF then
									Set RSIGPMHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
								end if
								
								if RSIGPMVencimento.EOF then
									Set RSIGPMVencimento = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & VencimentoDebito & "') AND TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
								end if
								
								if RSIGPMHoje.EOF then
									Set RSIGPMHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
								end if
								
								if RSIGPMVencimento.EOF then
									Set RSIGPMVencimento = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
								end if

								vValorCorrigido = (RSIGPMVencimento("COTA_Indice") / RSIGPMHoje("COTA_Indice")) * vValorCorrigido
								
								RSIGPMVencimento.Close
								RSIGPMHoje.Close
								Set RSIGPMVencimento = Nothing
								Set RSIGPMHoje = Nothing
							elseif vTaxaAtualizacao = 15 then
								'IGPM Pró-rata
								if CDate(AtualizaAte) > CDate(VencimentoDebito) then
									if Month(VencimentoDebito) = Month(AtualizaAte) and Year(VencimentoDebito) = Year(AtualizaAte) then
										Set RSIGPMVencimento = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & "01/" & Month(VencimentoDebito) & "/" & Year(VencimentoDebito) & "') AND TAXA_ID = " & vTaxaAtualizacao)
										if not RSIGPMVencimento.EOF then
											vValorCorrigido = DateDiff("d", VencimentoDebito, AtualizaAte) * RSIGPMVencimento("COTA_Indice") / 100 / 30 * vValorCorrigido
										end if
									else
										sai = false
										vTaxaJurosComposta = 0
										Set RSIGPMVencimento = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & "01/" & Month(VencimentoDebito) & "/" & Year(VencimentoDebito) & "') AND TAXA_ID = " & vTaxaAtualizacao)
										if not RSIGPMVencimento.EOF then
											vTaxaJurosComposta = 1 + Round((DateDiff("d", VencimentoDebito, DateAdd("m", 1, ("01/" & Month(VencimentoDebito) & "/" & Year(VencimentoDebito)))) * RSIGPMVencimento("COTA_Indice") / 100 / 30), 6)
										end if
										w = VencimentoDebito
										Do While not sai
											w = DateAdd("m", 1, w)
											if DateDiff("m", w, AtualizaAte) = 0 then
												sai = true
												Set RSIGPMVencimento = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & "01/" & Month(AtualizaAte) & "/" & Year(AtualizaAte) & "') AND TAXA_ID = " & vTaxaAtualizacao)
												if not RSIGPMVencimento.EOF then
													vTaxaJurosComposta = Round(vTaxaJurosComposta * Round((1 + (DateDiff("d", "01/" & Month(AtualizaAte) & "/" & Year(AtualizaAte), AtualizaAte) * RSIGPMVencimento("COTA_Indice") / 100 / 30)), 6), 6)
												end if
											else
												Set RSIGPMVencimento = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & "01/" & Month(w) & "/" & Year(w) & "') AND TAXA_ID = " & vTaxaAtualizacao)
												if not RSIGPMVencimento.EOF then
													vTaxaJurosComposta = Round(vTaxaJurosComposta * Round((1 + RSIGPMVencimento("COTA_Indice") / 100), 6), 6)
												end if
											end if
										Loop
										vValorCorrigido = vValorCorrigido + Round((vTaxaJurosComposta - 1) * vValorCorrigido, 2)
									end if
								end if
							else
								'outros
								Set RSIGPMHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = " & vTaxaAtualizacao)
								Set RSIGPMVencimento = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & VencimentoDebito & "') AND TAXA_ID = " & vTaxaAtualizacao)

								if RSIGPMHoje.EOF then
									Set RSIGPMHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
								end if
								
								if RSIGPMVencimento.EOF then
									Set RSIGPMVencimento = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & VencimentoDebito & "') AND TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
								end if
								
								if RSIGPMHoje.EOF then
									Set RSIGPMHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
								end if
								
								if RSIGPMVencimento.EOF then
									Set RSIGPMVencimento = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
								end if

								'if RSIGPMHoje.EOF or RSIGPMVencimento.EOF then
								'	vValorCorrigido = Valor
								'else
									vValorCorrigido = (RSIGPMHoje("COTA_Indice") / RSIGPMVencimento("COTA_Indice")) * vValorCorrigido
								'end if
								
								
								RSIGPMVencimento.Close
								RSIGPMHoje.Close
								Set RSIGPMVencimento = Nothing
								Set RSIGPMHoje = Nothing
							end if
						end if
						if vMulta then
							'Cobrar multa
							vValorMulta = 0
							if vMultaPrincipal then
								'Cobrar multa sobre o principal
								'vValorMulta = ((vNumDias * (vPercentMulta / 30)) / 100) * Valor
								vValorMulta = (vPercentMulta / 100) * Valor
							elseif vMultaPrincipalCorrigido then
								'Cobrar multa sobre o proncipal corrigido (com Juros)
								'vValorMulta = ((vNumDias * (vPercentMulta / 30)) / 100) * vValorCorrigido
								vValorMulta = (vPercentMulta / 100) * vValorCorrigido
							end if
							vValorCorrigido = vValorCorrigido + vValorMulta
						end if
					end if
				end if
			end if
		end if

		Conn.Close 
		Set Conn = Nothing
	end if
	
	ValorMulta = vValorMulta
end function


'Calcula o valor da correção de um título baseado na Politica de Negociação da empresa
function ValorMultaManual(ID_Carteira, VencimentoDebito, AtualizaAte, Valor, Sinal, PermiteAtualizacao, QtdParc, DataRecebimento, Multa, Juros, TDOC_ID, DataAtualizacao)
	Dim Conn
	Dim RSIGPMHoje
	Dim RSIGPMVencimento
	Dim RSPolitica
	Dim vAtualiza, vAtualizaDebito, vAtualizaCredito, vMulta, vPercentMulta, vMultaPrincipal, vMultaPrincipalCorrigido, vJuros, vTaxaJuros, vPercentJuros
	Dim vNumDias, vValorCorrigido, vValorMulta, vValorJuros, vValorHonorarios, vHonorarios, vPercentHonorarios, vHonorariosPrincipal, vHonorariosPrincipalCorrigido
	Dim vJurosSimples, vJurosComposto, vJurosFaixa, RSTaxaJurosFaixa, rsCont
	Dim vQtdMeses, sai, vDataNova
	Dim vVencimentoTitulo, vRecebimentoContrato, vAtualizaValor, vTaxaAtualizacao
	
	vQtdMeses = 0
	vValorCorrigido = Valor
	sai = false
	vValorCorrigido = Valor
	vValorJuros = 0
	vValorMulta = 0
	if PermiteAtualizacao then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open Application("Conexao")
	
		Set RSPolitica = Conn.Execute("SELECT * FROM Politica_de_Negociacao p WITH (NOLOCK) JOIN REL_PolNeg_Cart r ON p.PNEG_ID = r.PNEG_ID WHERE CART_ID = " & ID_Carteira)
		Set rsCont = Conn.Execute("SELECT CONT_ID FROM Carteiras WITH (NOLOCK) WHERE CART_ID = " & ID_Carteira)	
	
		if not RSPolitica.eof then
			vAtualiza = RSPolitica("PNEG_Atualiza")
			vAtualizaDebito = RSPolitica("PNEG_AtualizaDebito")
			vAtualizaCredito = RSPolitica("PNEG_AtualizaCredito")
			vMulta = RSPolitica("PNEG_Multa")
			vPercentMulta = Multa
			vMultaPrincipal = RSPolitica("PNEG_MultaPrincipal")
			vMultaPrincipalCorrigido = RSPolitica("PNEG_MultaPrincipalCorrigido")
			vAtualizaValor = RSPolitica("PNEG_AtualizaValor")
			vTaxaAtualizacao = RSPolitica("TAXA_ID_Atualizacao")
			vJuros = RSPolitica("PNEG_Juros")
			vJurosFaixa = RSPolitica("PNEG_JurosPorFaixa")
			vPercentJuros = Juros
			vJurosSimples = RSPolitica("PNEG_JurosSimples")
			vJurosComposto = RSPolitica("PNEG_JurosComposto")
			vVencimentoTitulo = RSPolitica("PNEG_AtualizaDoVencimento")
			vRecebimentoContrato = RSPolitica("PNEG_AtualizaDoRecebimento")
			
			'if rsCont("CONT_ID") = 3 then
			'	VencimentoDebito = DataVencDebito
			'end if

			if vRecebimentoContrato then
				vNumDias = DateDiff("d", DataRecebimento, AtualizaAte)
			else
				vNumDias = DateDiff("d", VencimentoDebito, AtualizaAte)
			end if
			
			Dim rsTaxaHoje, rsTaxaVenc
			
			if TDOC_ID = 67 and DateDiff("d", VencimentoDebito, AtualizaAte) < 61 then
				'Mora
				Set rsTaxaHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 4")
				Set rsTaxaVenc = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 4")

				if rsTaxaHoje.EOF then
					Set rsTaxaHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 4 ORDER BY COTA_Data DESC")
				end if
				
				if rsTaxaVenc.EOF then
					Set rsTaxaVenc = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 4 ORDER BY COTA_Data DESC")
				end if
				
				if not rsTaxaHoje.EOF and not rsTaxaVenc.EOF then
					vValorCorrigido = (rsTaxaHoje("COTA_Indice") / rsTaxaVenc("COTA_Indice")) * vValorCorrigido
				end if

				rsTaxaVenc.Close
				rsTaxaHoje.Close
				Set rsTaxaVenc = Nothing
				Set rsTaxaHoje = Nothing
			elseif TDOC_ID = 67 and DateDiff("d", VencimentoDebito, AtualizaAte) > 60 and DateDiff("d", VencimentoDebito, DataAtualizacao) < 61 then
				'Mora
				Set rsTaxaHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & DateAdd("d", 60, VencimentoDebito) & "') AND TAXA_ID = 4")
				Set rsTaxaVenc = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 4")

				if rsTaxaHoje.EOF then
					Set rsTaxaHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & DateAdd("d", 60, VencimentoDebito) & "') AND TAXA_ID = 4 ORDER BY COTA_Data DESC")
				end if
				
				if rsTaxaVenc.EOF then
					Set rsTaxaVenc = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 4 ORDER BY COTA_Data DESC")
				end if
				
				if not rsTaxaHoje.EOF and not rsTaxaVenc.EOF then
					vValorCorrigido = (rsTaxaHoje("COTA_Indice") / rsTaxaVenc("COTA_Indice")) * vValorCorrigido
				end if

				rsTaxaVenc.Close
				rsTaxaHoje.Close

				'CL/LP
				Set rsTaxaHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 5")
				Set rsTaxaVenc = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & DateAdd("d", 60, VencimentoDebito) & "') AND TAXA_ID = 5")

				if rsTaxaHoje.EOF then
					Set rsTaxaHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 5 ORDER BY COTA_Data DESC")
				end if
				
				if rsTaxaVenc.EOF then
					Set rsTaxaVenc = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & DateAdd("d", 60, VencimentoDebito) & "') AND TAXA_ID = 5 ORDER BY COTA_Data DESC")
				end if
				
				if not rsTaxaHoje.EOF and not rsTaxaVenc.EOF then
					vValorCorrigido = (rsTaxaHoje("COTA_Indice") / rsTaxaVenc("COTA_Indice")) * vValorCorrigido
				end if

				rsTaxaVenc.Close
				rsTaxaHoje.Close
				Set rsTaxaVenc = Nothing
				Set rsTaxaHoje = Nothing
				Set rsTaxaVenc = Nothing
				Set rsTaxaHoje = Nothing
			elseif (TDOC_ID = 67 and DateDiff("d", VencimentoDebito, AtualizaAte) > 60 and DateDiff("d", VencimentoDebito, DataAtualizacao) > 61) or TDOC_ID = 68 or TDOC_ID = 69 then
				'CL/LP
				Set rsTaxaHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 5")
				Set rsTaxaVenc = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 5")

				if rsTaxaHoje.EOF then
					Set rsTaxaHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 5 ORDER BY COTA_Data DESC")
				end if
				
				if rsTaxaVenc.EOF then
					Set rsTaxaVenc = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 5 ORDER BY COTA_Data DESC")
				end if
				
				if not rsTaxaHoje.EOF and not rsTaxaVenc.EOF then

					vValorCorrigido = (rsTaxaHoje("COTA_Indice") / rsTaxaVenc("COTA_Indice")) * vValorCorrigido
				end if

				rsTaxaVenc.Close
				rsTaxaHoje.Close
				Set rsTaxaVenc = Nothing
				Set rsTaxaHoje = Nothing
			else
				if vAtualiza then
					if (vAtualizaDebito and Sinal = "+") or (vAtualizaCredito and Sinal = "-") then
						if vJuros then
							vValorJuros = 0
							'Cobrar juros
							if vPercentJuros <> "" then
								if vJurosSimples then
									vValorJuros = (((vNumDias * (vPercentJuros / 30)) / 100) * vValorCorrigido)
								elseif vJurosComposto then
									vValorJuros = vValorCorrigido * ((vPercentJuros / 100))^(vNumDias / 30) 
								end if
							end if
						end if
						vValorCorrigido = vValorCorrigido + vValorJuros
						if vMulta then
							'Cobrar multa
							vValorMulta = 0
							if vMultaPrincipal then
								'Cobrar multa sobre o principal
								'vValorMulta = ((vNumDias * (vPercentMulta / 30)) / 100) * Valor
								vValorMulta = (vPercentMulta / 100) * Valor
							elseif vMultaPrincipalCorrigido then
								'Cobrar multa sobre o proncipal corrigido (com Juros)
								'vValorMulta = ((vNumDias * (vPercentMulta / 30)) / 100) * vValorCorrigido
								vValorMulta = (vPercentMulta / 100) * vValorCorrigido
							end if
							vValorCorrigido = vValorCorrigido + vValorMulta
						end if
					end if
				end if
			end if
		end if

		Conn.Close 
		Set Conn = Nothing
	end if
	
	ValorMultaManual = vValorMulta
end function

'Calcula o valor da correção de um título baseado na Politica de Negociação da empresa
function ValorCorrecaoManual(ID_Carteira, VencimentoDebito, AtualizaAte, Valor, Sinal, PermiteAtualizacao, QtdParc, DataRecebimento, PercentJuros, PercentMulta, TDOC_ID, DataAtualizacao)
	Dim Conn
	Dim RSIGPMHoje
	Dim RSIGPMVencimento
	Dim RSPolitica
	Dim vAtualiza, vAtualizaDebito, vAtualizaCredito, vMulta, vPercentMulta, vMultaPrincipal, vMultaPrincipalCorrigido, vJuros, vTaxaJuros, vPercentJuros
	Dim vNumDias, vValorCorrigido, vValorMulta, vValorJuros, vValorHonorarios, vHonorarios, vPercentHonorarios, vHonorariosPrincipal, vHonorariosPrincipalCorrigido
	Dim vJurosSimples, vJurosComposto, vJurosFaixa, RSTaxaJurosFaixa, rsCont
	Dim vQtdMeses, sai, vDataNova
	Dim vVencimentoTitulo, vRecebimentoContrato, vAtualizaValor, vTaxaAtualizacao
	
	vQtdMeses = 0
	vValorCorrigido = Valor
	vValorJuros = 0
	vValorMulta = 0
	if PermiteAtualizacao then
		Set Conn = Server.CreateObject("ADODB.Connection")
		Conn.Open Application("Conexao")
	
		Set RSPolitica = Conn.Execute("SELECT * FROM Politica_de_Negociacao p WITH (NOLOCK) JOIN REL_PolNeg_Cart r ON p.PNEG_ID = r.PNEG_ID WHERE CART_ID = " & ID_Carteira)
		Set rsCont = Conn.Execute("SELECT CONT_ID FROM Carteiras WITH (NOLOCK) WHERE CART_ID = " & ID_Carteira)	
	
		if not RSPolitica.eof then
			vAtualiza = RSPolitica("PNEG_Atualiza")
			vAtualizaDebito = RSPolitica("PNEG_AtualizaDebito")
			vAtualizaCredito = RSPolitica("PNEG_AtualizaCredito")
			vPercentMulta = PercentMulta
			vMultaPrincipal = RSPolitica("PNEG_MultaPrincipal")
			vMultaPrincipalCorrigido = RSPolitica("PNEG_MultaPrincipalCorrigido")
			vAtualizaValor = RSPolitica("PNEG_AtualizaValor")
			vTaxaAtualizacao = RSPolitica("TAXA_ID_Atualizacao")
			vPercentJuros = PercentJuros
			vJurosSimples = RSPolitica("PNEG_JurosSimples")
			vJurosComposto = RSPolitica("PNEG_JurosComposto")
			vVencimentoTitulo = RSPolitica("PNEG_AtualizaDoVencimento")
			vRecebimentoContrato = RSPolitica("PNEG_AtualizaDoRecebimento")
			
			if (ID_Carteira = 41 or ID_Carteira = 42 or ID_Carteira = 43 or ID_Carteira = 44) and AtualizaAte > CDate("17/10/2005") then
				AtualizaAte = CDate("17/10/2005")
			end if
			if vCONT_ID = 43 and AtualizaAte > CDate("10/04/2006") then
				AtualizaAte = CDate("10/04/2006")
			end if

			if vRecebimentoContrato then
				vNumDias = DateDiff("d", DataRecebimento, AtualizaAte)
			else
				vNumDias = DateDiff("d", VencimentoDebito, AtualizaAte)
			end if
			
			Dim rsTaxaHoje, rsTaxaVenc
			
			if TDOC_ID = 67 and DateDiff("d", VencimentoDebito, AtualizaAte) < 61 then
				'Mora
				Set rsTaxaHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 4")
				Set rsTaxaVenc = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 4")

				if rsTaxaHoje.EOF then
					Set rsTaxaHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 4 ORDER BY COTA_Data DESC")
				end if
				
				if rsTaxaVenc.EOF then
					Set rsTaxaVenc = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 4 ORDER BY COTA_Data DESC")
				end if
				
				if not rsTaxaHoje.EOF and not rsTaxaVenc.EOF then
					vValorCorrigido = (rsTaxaHoje("COTA_Indice") / rsTaxaVenc("COTA_Indice")) * vValorCorrigido
				end if

				rsTaxaVenc.Close
				rsTaxaHoje.Close
				Set rsTaxaVenc = Nothing
				Set rsTaxaHoje = Nothing
			elseif TDOC_ID = 67 and DateDiff("d", VencimentoDebito, AtualizaAte) > 60 and DateDiff("d", VencimentoDebito, DataAtualizacao) < 61 then
				'Mora
				Set rsTaxaHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & DateAdd("d", 60, VencimentoDebito) & "') AND TAXA_ID = 4")
				Set rsTaxaVenc = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 4")

				if rsTaxaHoje.EOF then
					Set rsTaxaHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & DateAdd("d", 60, VencimentoDebito) & "') AND TAXA_ID = 4 ORDER BY COTA_Data DESC")
				end if
				
				if rsTaxaVenc.EOF then
					Set rsTaxaVenc = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 4 ORDER BY COTA_Data DESC")
				end if
				
				if not rsTaxaHoje.EOF and not rsTaxaVenc.EOF then
					vValorCorrigido = (rsTaxaHoje("COTA_Indice") / rsTaxaVenc("COTA_Indice")) * vValorCorrigido
				end if

				rsTaxaVenc.Close
				rsTaxaHoje.Close

				'CL/LP
				Set rsTaxaHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 5")
				Set rsTaxaVenc = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & DateAdd("d", 60, VencimentoDebito) & "') AND TAXA_ID = 5")

				if rsTaxaHoje.EOF then
					Set rsTaxaHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 5 ORDER BY COTA_Data DESC")
				end if
				
				if rsTaxaVenc.EOF then
					Set rsTaxaVenc = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & DateAdd("d", 60, VencimentoDebito) & "') AND TAXA_ID = 5 ORDER BY COTA_Data DESC")
				end if
				
				if not rsTaxaHoje.EOF and not rsTaxaVenc.EOF then
					vValorCorrigido = (rsTaxaHoje("COTA_Indice") / rsTaxaVenc("COTA_Indice")) * vValorCorrigido
				end if

				rsTaxaVenc.Close
				rsTaxaHoje.Close
				Set rsTaxaVenc = Nothing
				Set rsTaxaHoje = Nothing
				Set rsTaxaVenc = Nothing
				Set rsTaxaHoje = Nothing
			elseif (TDOC_ID = 67 and DateDiff("d", VencimentoDebito, AtualizaAte) > 60 and DateDiff("d", VencimentoDebito, DataAtualizacao) > 61) or TDOC_ID = 68 or TDOC_ID = 69 then
				'CL/LP
				Set rsTaxaHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 5")
				Set rsTaxaVenc = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 5")

				if rsTaxaHoje.EOF then
					Set rsTaxaHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 5 ORDER BY COTA_Data DESC")
				end if
				
				if rsTaxaVenc.EOF then
					Set rsTaxaVenc = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 5 ORDER BY COTA_Data DESC")
				end if
				
				if not rsTaxaHoje.EOF and not rsTaxaVenc.EOF then
					vValorCorrigido = (rsTaxaHoje("COTA_Indice") / rsTaxaVenc("COTA_Indice")) * vValorCorrigido
				end if

				rsTaxaVenc.Close
				rsTaxaHoje.Close
				Set rsTaxaVenc = Nothing
				Set rsTaxaHoje = Nothing
			else
				if vAtualiza then
					if (vAtualizaDebito and Sinal = "+") or (vAtualizaCredito and Sinal = "-") then
						vValorJuros = 0
						'Cobrar juros
						if vPercentJuros <> "" then
							if vJurosSimples then
								vValorJuros = (((vNumDias * (vPercentJuros / 30)) / 100) * vValorCorrigido)
							elseif vJurosComposto then
								vValorJuros = vValorCorrigido * ((vPercentJuros / 100))^(vNumDias / 30) 
							end if
						end if
						if vAtualizaValor then
							if vTaxaAtualizacao = 3 then
								'IGPM
								Set RSIGPMHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 3")
								Set RSIGPMVencimento = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & VencimentoDebito & "') + 2 AND TAXA_ID = 3")

								if RSIGPMHoje.EOF then
									Set RSIGPMHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 3 ORDER BY COTA_Data DESC")
								end if
								
								if RSIGPMVencimento.EOF then
									Set RSIGPMVencimento = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & VencimentoDebito & "') AND TAXA_ID = 3 ORDER BY COTA_Data DESC")
								end if
								
								vValorCorrigido = (RSIGPMHoje("COTA_Indice") / RSIGPMVencimento("COTA_Indice")) * vValorCorrigido
								
								RSIGPMVencimento.Close
								RSIGPMHoje.Close
								Set RSIGPMVencimento = Nothing
								Set RSIGPMHoje = Nothing
							elseif vTaxaAtualizacao = 10 then
								'Taxa Renner
								Set RSIGPMHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = " & vTaxaAtualizacao)
								Set RSIGPMVencimento = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & VencimentoDebito & "') AND TAXA_ID = " & vTaxaAtualizacao)

								if RSIGPMHoje.EOF then
									Set RSIGPMHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
								end if
								
								if RSIGPMVencimento.EOF then
									Set RSIGPMVencimento = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & VencimentoDebito & "') AND TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
								end if
								
								vValorCorrigido = (RSIGPMVencimento("COTA_Indice") / RSIGPMHoje("COTA_Indice")) * vValorCorrigido
								
								
								RSIGPMVencimento.Close
								RSIGPMHoje.Close
								Set RSIGPMVencimento = Nothing
								Set RSIGPMHoje = Nothing
							else
								'outros
								Set RSIGPMHoje = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = " & vTaxaAtualizacao)
								Set RSIGPMVencimento = Conn.Execute("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & VencimentoDebito & "') AND TAXA_ID = " & vTaxaAtualizacao)

								if RSIGPMHoje.EOF then
									Set RSIGPMHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
								end if
								
								if RSIGPMVencimento.EOF then
									Set RSIGPMVencimento = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & VencimentoDebito & "') AND TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
								end if
								
								if RSIGPMHoje.EOF then
									Set RSIGPMHoje = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
								end if
								
								if RSIGPMVencimento.EOF then
									Set RSIGPMVencimento = Conn.Execute("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
								end if

								'if RSIGPMHoje.EOF or RSIGPMVencimento.EOF then
								'	vValorCorrigido = Valor
								'else
									vValorCorrigido = (RSIGPMHoje("COTA_Indice") / RSIGPMVencimento("COTA_Indice")) * vValorCorrigido
								'end if
								
								
								RSIGPMVencimento.Close
								RSIGPMHoje.Close
								Set RSIGPMVencimento = Nothing
								Set RSIGPMHoje = Nothing
							end if
						end if
						vValorCorrigido = vValorCorrigido + vValorJuros
						'Cobrar multa
						vValorMulta = 0
						if vMultaPrincipal then
							'Cobrar multa sobre o principal
							vValorMulta = (vPercentMulta / 100) * Valor
						elseif vMultaPrincipalCorrigido then
							'Cobrar multa sobre o proncipal corrigido (com Juros)
							vValorMulta = (vPercentMulta / 100) * vValorCorrigido
						end if
						vValorCorrigido = vValorCorrigido + vValorMulta
					end if
				end if
			end if
		end if

		Conn.Close 
		Set Conn = Nothing
	end if
	
	ValorCorrecaoManual = vValorCorrigido - Valor
end function

'Rotina de data da Nota Promissória
Function DataTextoNP(fDia, fMes, fAno)
	Dim vDia, vMes, vAno, Dia, Mes, Ano
	Dia = CStr(fDia)
	Mes = CInt(fMes)
	Ano = CStr(fAno)
	Dia = Right("0" & Dia, 2) '00
	Ano = Right("200" & Ano, 4) '0000
	'Primeira casa do Dia
	if Mid(Dia, 1, 1) = "0" then
		vDia = ""
	elseif Mid(Dia, 1, 1) = "1" then
		vDia = "Décimo "
	elseif Mid(Dia, 1, 1) = "2" then
		vDia = "Vigésimo "
	elseif Mid(Dia, 1, 1) = "3" then
		vDia = "Trigésimo "
	end if
	'Segunda casa do Dia
	if Mid(Dia, 2, 1) = "1" then
		vDia = vDia & "Primeiro"
	elseif Mid(Dia, 2, 1) = "2" then
		vDia = vDia & "Segundo"
	elseif Mid(Dia, 2, 1) = "3" then
		vDia = vDia & "Terceiro"
	elseif Mid(Dia, 2, 1) = "4" then
		vDia = vDia & "Quarto"
	elseif Mid(Dia, 2, 1) = "5" then
		vDia = vDia & "Quinto"
	elseif Mid(Dia, 2, 1) = "6" then
		vDia = vDia & "Sexto"
	elseif Mid(Dia, 2, 1) = "7" then
		vDia = vDia & "Sétimo"
	elseif Mid(Dia, 2, 1) = "8" then
		vDia = vDia & "Oitavo"
	elseif Mid(Dia, 2, 1) = "9" then
		vDia = vDia & "Nono"
	end if
	'Mês
	if Mes = 1 then
		vMes = "Janeiro"
	elseif Mes = 2 then
		vMes = "Fevereiro"
	elseif Mes = 3 then
		vMes = "Março"
	elseif Mes = 4 then
		vMes = "Abril"
	elseif Mes = 5 then
		vMes = "Maio"
	elseif Mes = 6 then
		vMes = "Junho"
	elseif Mes = 7 then
		vMes = "Julho"
	elseif Mes = 8 then
		vMes = "Agosto"
	elseif Mes = 9 then
		vMes = "Setembro"
	elseif Mes = 10 then
		vMes = "Outubro"
	elseif Mes = 11 then
		vMes = "Novembro"
	elseif Mes = 12 then
		vMes = "Dezembro"
	end if
	'Ano
	if Mid(Ano, 3, 2) = "00" then
		vAno = "Dois Mil"
	elseif Mid(Ano, 3, 2) = "01" then
		vAno = "Dois Mil e Um"
	elseif Mid(Ano, 3, 2) = "02" then
		vAno = "Dois Mil e Dois"
	elseif Mid(Ano, 3, 2) = "03" then
		vAno = "Dois Mil e Três"
	elseif Mid(Ano, 3, 2) = "04" then
		vAno = "Dois Mil e Quatro"
	elseif Mid(Ano, 3, 2) = "05" then
		vAno = "Dois Mil e Cinco"
	elseif Mid(Ano, 3, 2) = "06" then
		vAno = "Dois Mil e Seis"
	elseif Mid(Ano, 3, 2) = "07" then
		vAno = "Dois Mil e Sete"
	elseif Mid(Ano, 3, 2) = "08" then
		vAno = "Dois Mil e Oito"
	elseif Mid(Ano, 3, 2) = "09" then
		vAno = "Dois Mil e Nove"
	elseif Mid(Ano, 3, 2) = "10" then
		vAno = "Dois Mil e Dez"
	elseif Mid(Ano, 3, 2) = "11" then
		vAno = "Dois Mil e Onze"
	elseif Mid(Ano, 3, 2) = "12" then
		vAno = "Dois Mil e Doze"
	elseif Mid(Ano, 3, 2) = "13" then
		vAno = "Dois Mil e Treze"
	elseif Mid(Ano, 3, 2) = "14" then
		vAno = "Dois Mil e Quatorze"
	elseif Mid(Ano, 3, 2) = "15" then
		vAno = "Dois Mil e Quinze"
	elseif Mid(Ano, 3, 2) = "16" then
		vAno = "Dois Mil e Dezesseis"
	elseif Mid(Ano, 3, 2) = "17" then
		vAno = "Dois Mil e Dezessete"
	elseif Mid(Ano, 3, 2) = "18" then
		vAno = "Dois Mil e Dezoito"
	elseif Mid(Ano, 3, 2) = "19" then
		vAno = "Dois Mil e Dezenove"
	elseif Mid(Ano, 3, 2) = "20" then
		vAno = "Dois Mil e Vinte"
	elseif Mid(Ano, 3, 2) = "21" then
		vAno = "Dois Mil e Vinte e Um"
	elseif Mid(Ano, 3, 2) = "22" then
		vAno = "Dois Mil e Vinte e Dois"
	elseif Mid(Ano, 3, 2) = "23" then
		vAno = "Dois Mil e Vinte e Três"
	elseif Mid(Ano, 3, 2) = "24" then
		vAno = "Dois Mil e Vinte e Quatro"
	elseif Mid(Ano, 3, 2) = "25" then
		vAno = "Dois Mil e Vinte e Cinco"
	elseif Mid(Ano, 3, 2) = "26" then
		vAno = "Dois Mil e Vinte e Seis"
	elseif Mid(Ano, 3, 2) = "27" then
		vAno = "Dois Mil e Vinte e Sete"
	elseif Mid(Ano, 3, 2) = "28" then
		vAno = "Dois Mil e Vinte e Oito"
	elseif Mid(Ano, 3, 2) = "29" then
		vAno = "Dois Mil e Vinte e Nove"
	elseif Mid(Ano, 3, 2) = "30" then
		vAno = "Dois Mil e Trinta"
	end if
	DataTextoNP = vDia & " Dia do Mês de " & vMes & " de " & vAno
End Function

'Rotina de Extenso
Dim x_Centavos, x_I, x_J, x_K, x_Numero, x_QtdCentenas, x_TotCentenas, x_TxtExtenso( 900 ) 
Dim x_TxtMoeda( 6 ), x_ValCentena( 6 ), x_Valor, x_ValSoma

' Matrizes de textos
x_TxtMoeda( 1 ) = "rea"
x_TxtMoeda( 2 ) = "mil"
x_TxtMoeda( 3 ) = "milh"
x_TxtMoeda( 4 ) = "bilh"
x_TxtMoeda( 5 ) = "trilh"

x_TxtExtenso( 1 ) = "um"
x_TxtExtenso( 2 ) = "dois"
x_TxtExtenso( 3 ) = "tres"
x_TxtExtenso( 4 ) = "quatro"
x_TxtExtenso( 5 ) = "cinco"
x_TxtExtenso( 6 ) = "seis"
x_TxtExtenso( 7 ) = "sete"
x_TxtExtenso( 8 ) = "oito"
x_TxtExtenso( 9 ) = "nove"
x_TxtExtenso( 10 ) = "dez"
x_TxtExtenso( 11 ) = "onze"
x_TxtExtenso( 12 ) = "doze"
x_TxtExtenso( 13 ) = "treze"
x_TxtExtenso( 14 ) = "quatorze"
x_TxtExtenso( 15 ) = "quinze"
x_TxtExtenso( 16 ) = "dezesseis"
x_TxtExtenso( 17 ) = "dezessete"
x_TxtExtenso( 18 ) = "dezoito"
x_TxtExtenso( 19 ) = "dezenove"
x_TxtExtenso( 20 ) = "vinte"
x_TxtExtenso( 30 ) = "trinta"
x_TxtExtenso( 40 ) = "quarenta"
x_TxtExtenso( 50 ) = "cinquenta"
x_TxtExtenso( 60 ) = "sessenta"
x_TxtExtenso( 70 ) = "setenta"
x_TxtExtenso( 80 ) = "oitenta"
x_TxtExtenso( 90 ) = "noventa"
x_TxtExtenso( 100 ) = "cento"
x_TxtExtenso( 200 ) = "duzentos"
x_TxtExtenso( 300 ) = "trezentos"
x_TxtExtenso( 400 ) = "quatrocentos"
x_TxtExtenso( 500 ) = "quinhentos"
x_TxtExtenso( 600 ) = "seiscentos"
x_TxtExtenso( 700 ) = "setecentos"
x_TxtExtenso( 800 ) = "oitocentos"
x_TxtExtenso( 900 ) = "novecentos"

' Função Principal de Conversão de Valores em Extenso
Function Extenso( x_Numero )

	x_Numero = FormatNumber( x_Numero , 2 )
	x_Centavos = right( x_Numero , 2 )
	x_ValCentena( 0 ) = 0
	x_QtdCentenas = int( ( len( x_Numero ) + 1 ) / 4 )

	For x_I = 1 to x_QtdCentenas
		x_ValCentena( x_I ) = "" 
	Next
	'
	x_I = 1
	x_J = 1
	For x_K = len( x_Numero ) - 3 to 1 step -1
		x_ValCentena( x_J ) = mid( x_Numero , x_K , 1 ) & x_ValCentena( x_J )
		if x_I / 3 = int( x_I / 3 ) then
			x_J = x_J + 1
			x_K = x_K - 1
		end if
		x_I = x_I + 1
	next
	x_TotCentenas = 0
	Extenso = "" 
	For x_I = x_QtdCentenas to 1 step -1

		x_TotCentenas = x_TotCentenas + int( x_ValCentena( x_I ) )

		if int( x_ValCentena( x_I ) ) <> 0 or ( int( x_ValCentena( x_I ) ) = 0 and x_I = 1 )then
			if int( x_ValCentena( x_I ) = 0 and int( x_ValCentena( x_I + 1 ) ) = 0 and x_I = 1 )then
				Extenso = Extenso & ExtCentena( x_ValCentena( x_I ) , x_TotCentenas ) & " de " & x_TxtMoeda( x_I )
			else
				Extenso = Extenso & ExtCentena( x_ValCentena( x_I ) , x_TotCentenas ) & " " & x_TxtMoeda( x_I )
			end if
			if int( x_ValCentena( x_I ) ) <> 1 or ( x_I = 1 and x_TotCentenas <> 1 ) then
				Select Case x_I
				Case 1
					Extenso = Extenso & "is"
				Case 3, 4, 5
					Extenso = Extenso & "ões"
				End Select 
			else
				Select Case x_I
				Case 1
					Extenso = Extenso & "l"
				Case 3, 4, 5
					Extenso = Extenso & "ão"
				End Select 
			end if
		end if
		if int( x_ValCentena( x_I - 1 ) ) = 0 then
			Extenso = Extenso
		else
			if ( int( x_ValCentena( x_I + 1 ) ) = 0 and ( x_I + 1 ) <= x_QtdCentenas ) or ( x_I = 2 ) then
				Extenso = Extenso & " e "
			else
				Extenso = Extenso & ", "
			end if
		end if 
	next

	if x_Centavos > 0 then
		if int( x_Centavos ) = 1 then
			Extenso = Extenso & " e " & ExtDezena( x_Centavos ) & " centavo"
		else
			Extenso = Extenso & " e " & ExtDezena( x_Centavos ) & " centavos"
		end if
	end if
	Extenso = UCase( Left( Extenso , 1 ) )&right( Extenso , len( Extenso ) - 1 )
End Function

Function ExtDezena( x_Valor )
	' Retorna o Valor por Extenso referente à DEZENA recebida
	ExtDezena = ""
	if int( x_Valor ) > 0 then
		if int( x_Valor ) < 20 then
			ExtDezena = x_TxtExtenso( int( x_Valor ) )
		else
			ExtDezena = x_TxtExtenso( int( int( x_Valor ) / 10 ) * 10 )
			if ( int( x_Valor ) / 10 ) - int( int( x_Valor ) / 10 ) <> 0 then
				ExtDezena = ExtDezena & " e " & x_TxtExtenso( int( right( x_Valor , 1 ) ) )
			end if
		end if
	end if
End Function

Function ExtCentena( x_Valor, x_ValSoma )
	ExtCentena = ""

	if int( x_Valor ) > 0 then

		if int( x_Valor ) = 100 then
			ExtCentena = "cem"
		else
			if int( x_Valor ) < 20 then
				if int( x_Valor ) = 1 then
					If x_ValSoma - int( x_Valor ) = 0 then
						ExtCentena = "hum"
					else
						ExtCentena = x_TxtExtenso( int( x_Valor ) )
					end if
				else
					ExtCentena = x_TxtExtenso( int( x_Valor ) )
				end if
			else
				if int( x_Valor ) < 100 then
					ExtCentena = ExtDezena( right( x_Valor , 2 ) )
				else 
					ExtCentena = x_TxtExtenso( int( int( x_Valor ) / 100 )*100 )
					if ( int( x_Valor ) / 100 ) - int( int( x_Valor ) / 100 ) <> 0 then
						ExtCentena = ExtCentena & " e " & ExtDezena( right( x_Valor , 2 ) )
					end if
				end if
			end if
		end if
	end if
End Function

Function Obrigadorio()
	Response.Write "<font Color='#FF0000'>*</Font>"
End Function

Function ExibirObrigadorio()
	Response.Write "<font Color='#FF0000' size=2>*</Font><font size=1>Campos com preenchimento obrigatório</font>"
End Function

Function Aguarde_INI()
	Response.Write	"<div id=aguarde style='position:absolute; width:100%; left: 0px; top: 0px; overflow: auto;'>" & vbCrLf & _
					"<br>" &_
					"<br>" &_
					"<br>" &_
					"<br>" &_
					"<table border=0 cellspacing=0 cellpadding=0 align=center ID=Table1>" & vbCrLf & _
					"<tr>" & vbCrLf & _
					"<td valign=top class=texto1><img src='../../images/status_wait.gif'></td>" & vbCrLf & _
					"</tr>" & vbCrLf & _
					"</table>" & vbCrLf & _
					"</div>" & vbCrLf   
	Response.Flush 
End Function

Function Aguarde_FIM()
	Response.Write	"<script language=javascript>" & vbCrLf & _
					"document.getElementById(""aguarde"").style.visibility = ""hidden""" & vbCrLf & _
					"</script>" & vbCrLf 
End Function

Function RetiraCaracteresEspeciais(str)
	Dim Tira, caracteres, carac, i 
	Tira = Trim(str)
'	caracteres = "! # $ % ^ & * ( ) = + { } [ ] | ; : / ? > , < ' ¨ ´ `"
	caracteres = "'"
	carac = Split(caracteres," ")
	for i = LBound(carac) to UBound(carac)
		Tira = replace(Tira,carac(i),"")
	next
	RetiraCaracteresEspeciais = Tira
End Function

function Formatohhmmss(seg_total)
	dim seg, min, hora
	min = int(seg_total / 60)
	seg = int(seg_total - (min * 60))
	hora = int(min / 60)
	min = min - (hora * 60)
	if  min < 10 then min = "0" & min
	if  seg < 10 then seg = "0" & seg
	if  hora < 10 then hora = "0" & hora
	 
	Formatohhmmss = hora & ":" & min & ":" & seg
end function

Function FinalizarRel()
	Dim TempoFinalRel
	TempoFinalRel = Datediff("S",HoraIniRel9586325,now)	
	IF TempoFinalRel = 0 then 
		TempoFinalRel = "menos de 1 segundo"
	Else
		TempoFinalRel = Formatohhmmss(TempoFinalRel)
	End if 
	Response.Write "<Br>" & _
				   "<table cellpadding=0 cellspacing=0 width=100% >" & _
						"<tr>" & _
							"<td align=right><font face=verdana size=1>Relatório processado em " & TempoFinalRel & "</font></td>" & _
						"</tr>" & _
					"</table>"
End Function
%>
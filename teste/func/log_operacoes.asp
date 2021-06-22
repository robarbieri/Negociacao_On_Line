<%
Function TiraAspas(Texto)
	TiraAspas = replace(Texto,"'","''")
End Function

Function RetornaProximoNumero(Tabela, Campo, Filial)
	'Dim BD
	'Dim RS

	'Set BD = Server.CreateObject("ADODB.Connection")
	'BD.Open Application("Conexao")

	'Set RS = BD.Execute("SELECT MAX(" & Campo & ") + 1000 AS " & Campo & " FROM " & Tabela & " WITH (NOLOCK) WHERE " & Campo & " LIKE '%' + RIGHT('00' + Convert(varchar, " & Filial & "), 3)")

	'If IsNull(RS(Campo)) or RS(Campo) = "" then
	'	RetornaProximoNumero = 1000 + Filial
	'Else
	'  	RetornaProximoNumero = RS(Campo)
	'End if

	'RS.Close
	'Set RS = Nothing
	'BD.Close
	'Set BD = Nothing

	Dim BD
	Dim RS
	Dim vID

	Set BD = Server.CreateObject("ADODB.Connection")
	BD.Open Application("Conexao")

	Set RS = BD.Execute("SELECT NEWID()")
	vID = RS(0)
	RS.Close
	
	'BD.Execute("INSERT INTO Identificadores " &_
	'			"SELECT '" & Tabela & "' AS IDEN_Tabela, '" & Campo & "' AS IDEN_Campo, IDEN_Valor = CASE WHEN MAX(IDEN_Valor) IS NULL THEN 1000 + " & Filial & " ELSE MAX(IDEN_Valor) END, '" & vID & "' AS IDEN_Codigo, GETDATE() AS IDEN_DataHora " &_
	'			"FROM  " &_
	'			"(SELECT MAX(" & Campo & ") + 1000 AS IDEN_Valor FROM " & Tabela & " WITH (NOLOCK) " &_
	'			"WHERE " & Campo & " LIKE '%' + RIGHT('00' + Convert(varchar, " & Filial & "), 3) " &_
	'			"UNION " &_
	'			"SELECT MAX(IDEN_Valor) + 1000 AS IDEN_Valor FROM Identificadores WITH (NOLOCK) " &_
	'			"WHERE IDEN_Tabela = '" & Tabela & "' AND IDEN_Campo = '" & Campo & "' AND IDEN_Valor LIKE '%' + RIGHT('00' + Convert(varchar, " & Filial & "), 3)) r ")
	
	Set RS = BD.Execute("SELECT * FROM Identificadores WITH (NOLOCK) WHERE IDEN_Codigo = '" & vID & "'")

	vID = RS("IDEN_Valor")

	RS.Close
	Set RS = Nothing
	BD.Close
	Set BD = Nothing

	RetornaProximoNumero = vID
End Function


Sub PreencheLOG(Comando, FilialDestino)
	Dim FILI_ID_Origem
	Dim FILI_ID_Destino
	Dim BD
	Dim Cmd

	If Application("PreencheLOG") = "S" then
		Set BD = Server.CreateObject("ADODB.Connection")
		BD.Open Application("Conexao")

		FILI_ID_Origem = Session("FILI_ID")
		If FilialDestino = -1 Or FilialDestino = "-1" Then
			'O destino são todas as filiais
			FILI_ID_Destino = "NULL"
		Else
			'O destino é uma filial específica
			FILI_ID_Destino = FilialDestino
		End If
	    
		Cmd = TiraAspas(Comando)
		'BD.Execute "EXEC prc_CriaLogOperacao @LOOP_Comando = '" & Cmd & "', @FILI_ID_Origem = " & FILI_ID_Origem & ", @FILI_ID_Destino = " & FILI_ID_Destino

		BD.Close
		Set BD = Nothing
	End If
End Sub

Sub PreencheLOG_Contratante(Comando, FilialDestino, CONT_ID)
	Dim FILI_ID_Origem
	Dim FILI_ID_Destino
	Dim BD
	Dim Cmd

	If Application("PreencheLOG") = "S" then
		Set BD = Server.CreateObject("ADODB.Connection")
		BD.Open Application("Conexao")

		FILI_ID_Origem = Session("FILI_ID")
		If FilialDestino = -1 Or FilialDestino = "-1" Then
			'O destino são todas as filiais
			FILI_ID_Destino = "NULL"
		Else
			'O destino é uma filial específica
			FILI_ID_Destino = FilialDestino
		End If
	    
		Cmd = TiraAspas(Comando)
		'BD.Execute "EXEC prc_INS_LogOperacaoContratante @LOOP_Comando = '" & Cmd & "', @FILI_ID_Origem = " & FILI_ID_Origem & ", @FILI_ID_Destino = " & FILI_ID_Destino & ", @CONT_ID = " & CONT_ID

		BD.Close
		Set BD = Nothing
	End If
End Sub

Sub PreencheLOG_Acordo(Comando, FilialDestino, CONT_ID, ACOR_ID)
	Dim FILI_ID_Origem
	Dim FILI_ID_Destino
	Dim BD
	Dim Cmd
	Dim vData, rsAcordo
	
	'
	'Verificar se o acordo está no loop e ainda não foi enviado, se estiver, usar a mesma data, senão, 
	'usar a data do dia
	'
	
	If Application("PreencheLOG") = "S" then
		Set BD = Server.CreateObject("ADODB.Connection")
		BD.Open Application("Conexao")

		Set rsAcordo = bd.execute("SELECT * FROM Log_de_Operacoes WHERE ACOR_ID = " & ACOR_ID & " AND LOOP_Transmitido = 0 ORDER BY LOOP_ID")
		if rsAcordo.EOF then
			vData = date
		else
			vData = rsAcordo("LOOP_DataParaEnvio")
		end if
		rsAcordo.Close
		Set rsAcordo = Nothing

		FILI_ID_Origem = Session("FILI_ID")
		If FilialDestino = -1 Or FilialDestino = "-1" Then
			'O destino são todas as filiais
			FILI_ID_Destino = "NULL"
		Else
			'O destino é uma filial específica
			FILI_ID_Destino = FilialDestino
		End If
	    
		Cmd = TiraAspas(Comando)
		'BD.Execute "EXEC prc_INS_LogOperacaoAcordos @LOOP_Comando = '" & Cmd & "', @FILI_ID_Origem = " & FILI_ID_Origem & ", @FILI_ID_Destino = " & FILI_ID_Destino & ", @CONT_ID = " & CONT_ID & ", @ACOR_ID = " & ACOR_ID & ", @LOOP_DataParaEnvio = '" & vData & "', @LOOP_Acordo = 0"

		BD.Close
		Set BD = Nothing
	End If
End Sub

Sub PreencheLOG_Acordo2(Comando, FilialDestino, CONT_ID, ACOR_ID, Acordo)
	Dim FILI_ID_Origem
	Dim FILI_ID_Destino
	Dim BD
	Dim Cmd
	Dim vData, rsAcordo
	
	'
	'Verificar se o acordo está no loop e ainda não foi enviado, se estiver, usar a mesma data, senão, 
	'usar a data do dia
	'
	
	If Application("PreencheLOG") = "S" then
		Set BD = Server.CreateObject("ADODB.Connection")
		BD.Open Application("Conexao")

		'Set rsAcordo = bd.execute("SELECT * FROM Log_de_Operacoes WHERE ACOR_ID = " & ACOR_ID & " AND LOOP_Transmitido = 0 ORDER BY LOOP_ID")
		'if rsAcordo.EOF then
			vData = date
		'else
		'	vData = rsAcordo("LOOP_DataParaEnvio")
		'end if
		'rsAcordo.Close
		'Set rsAcordo = Nothing

		FILI_ID_Origem = Session("FILI_ID")
		If FilialDestino = -1 Or FilialDestino = "-1" Then
			'O destino são todas as filiais
			FILI_ID_Destino = "NULL"
		Else
			'O destino é uma filial específica
			FILI_ID_Destino = FilialDestino
		End If
	    
		Cmd = TiraAspas(Comando)
		'BD.Execute "EXEC prc_INS_LogOperacaoAcordos @LOOP_Comando = '" & Cmd & "', @FILI_ID_Origem = " & FILI_ID_Origem & ", @FILI_ID_Destino = " & FILI_ID_Destino & ", @CONT_ID = " & CONT_ID & ", @ACOR_ID = " & ACOR_ID & ", @LOOP_DataParaEnvio = '" & vData & "', @LOOP_Acordo = " & Acordo

		BD.Close
		Set BD = Nothing
	End If
End Sub
%>

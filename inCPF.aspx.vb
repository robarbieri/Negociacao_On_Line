Option Explicit On
Option Strict Off

Imports ConnectTo
Imports System.Data.SqlClient
Imports XMail

Partial Class inCPF
    Inherits System.Web.UI.Page

    Protected eMail As New XMail.SendMail

    Protected Sub btnCPF_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCPF.Click

        Dim Mirror As New Comando
        Dim objDR As SqlDataReader = Nothing
        Dim objDRDev As SqlDataReader = Nothing
        Dim objDRAC As SqlDataReader = Nothing
        Dim strSql As String = ""
        Dim strCPF As String = ""
        Dim strNascimento As String = ""
        Dim strContratos As String = ""
        Dim x As Integer = 0
        Dim Funcoes As New Funcoes.Local

        'Try

        btnCPF.Enabled = False

        If Trim(txtCPF.Text) = "" Or Trim(txtDataNasc.Text) = "" Then
            ShowMessage("Preencha os campos CPF e Data de Nascimento.")
            'Exit Try
            btnCPF.Enabled = True
            Exit Sub
        End If

        If Trim(txtCodigo.Text) = "" Then txtCodigo.Text = "CITI"

        strCPF = Trim(Funcoes.SoNumeros(txtCPF.Text))
        strNascimento = Trim(txtDataNasc.Text)

        Session("Saldo") = ""
        Session("CPF") = strCPF
        Session("Nascimento") = strNascimento
        Session("Codigo") = UpperTrim(txtCodigo.Text)

        Mirror.Banco = "NEOWEB"
        'Mirror.Banco = "NEOWEB_REAL"

        strSql = "Select A.CTRA_Numero, B.DEVE_Nome Nome From Contratos A (NOLOCK) " & _
                 "Join Devedores B (NOLOCK) On B.DEVE_Id = A.DEVE_Id " & _
                 "Join Status_do_Contrato C (NOLOCK) On A.SCON_Id = C.SCON_Id " & _
                 "Join Carteiras D (NOLOCK) On A.CART_Id = D.CART_Id " & _
                 "Join Contratante E (NOLOCK) On D.CONT_Id = E.CONT_Id " & _
                 "Where E.CONT_Id = 6 AND C.SCON_Id Not In(4,7) AND B.DEVE_CGCCPF = '" & strCPF & "' " & _
                 "AND B.DEVE_Nascimento = '" & Format(CDate(strNascimento), "yyy-MM-dd") & "'"
        objDR = Mirror.ExecuteQuery(strSql)

        Do While objDR.Read

            Session("Nome") = Trim(objDR("Nome").ToString)

            strSql = "Select TOP 1 1 From Contratos Where CTRA_Numero = '" & Trim(objDR(0).ToString) & "' " & _
                     "AND DateDiff(day,GetDate(),CTRA_DataDevolucao) > 15" 'VEFIRICA DATA DE DEVOLUCAO - TRAVAR 15 DIAS ANTES DA DEVOLUCAO
            objDRDev = Mirror.ExecuteQuery(strSql)

            If objDRDev.Read Then

                If x > 0 Then strContratos = strContratos & ","
                strContratos = strContratos & Trim(objDR(0).ToString)

                x = x + 1

            End If

        Loop

        objDR.Close()

        If Trim(strContratos) = "" Then
            ShowMessage("Não foi localizado um contrato ativo para este CPF. Verifique se os dados estão corretos. Caso estejam, entre em contato com nosso atendimento de Seg à Sex, das 9:00 às 19:00, no telefone SP 11-2171-0301, demais localidades 0800-6000-500.")
            btnCPF.Enabled = True
            'Exit Try
            Exit Sub
        End If

        Mirror.Banco = "MIRRORWEB"
        strSql = "Select TOP 1 1 from tb_Faixas_de_NegociacaoOnLine Where FNEG_Campanha = '" & UpperTrim(txtCodigo.Text) & "'"
        objDR = Mirror.ExecuteQuery(strSql)

        If Not objDR.Read Then
            ShowMessage("Código de campanha " & UpperTrim(txtCodigo.Text) & " inválido, verifique o valor digitado.")
            'Exit Try
            btnCPF.Enabled = True
            Exit Sub
        End If

        If x > 1 Then
            objDR.Close()
            Response.Redirect("inContrato.aspx?ctra=" & Trim(strContratos) & "&cod=" & UpperTrim(txtCodigo.Text), False)
        Else
            Mirror.Banco = "MIRRORWEB"
            strSql = "prc_INS_NEGLOGIN '" & SoNumeros(Session("CPF")) & "','" & SoNumeros(Trim(strContratos)) & "'"
            Mirror.Execute(strSql)

            strSql = "prc_SEL_NEGLOGIN '" & SoNumeros(Trim(strContratos)) & "'"
            objDR = Mirror.ExecuteQuery(strSql)

            If Not objDR.Read Then
                SendMailErro("Erro na Negociação On-Line." & Chr(13) & Chr(13) & "Número: " & Err.Number & Chr(13) & "Source: " & Err.Source & Chr(13) & "Descrição: " & Err.Description & Chr(13) & "Help: " & Chr(13) & Err.HelpFile & Chr(13) & Err.HelpContext & Chr(13) & Err.LastDllError)
                ShowMessage("Sistema Inativo. Por favor tente novamente em alguns minutos ou entre em contato com nosso atendimento no telefone SP 11-2171-0301, demais localidades 0800-6000-500.")
                'Response.Redirect("inCPF.aspx")
                'Exit Try
                btnCPF.Enabled = True
                objDR.Close()
                Exit Sub
            End If

            Session("Contrato") = Trim(strContratos)
            Session("ID") = Trim(objDR("ID").ToString)

            GravarInput(Trim(strContratos), "<b>Acesso à Negociação On-Line.<br>ID: " & Trim(objDR("ID").ToString) & "<br>Data: " & Format(CDate(Now), "dd/MM/yyy HH:mm:ss") & "</b>", "N_ONLINE")

            strSql = "prc_SEL_AtualizacaoCadastral " & SoNumeros(Session("CPF"))
            objDRAC = Mirror.ExecuteQuery(strSql)

            If Not objDRAC.Read Then
                'Response.Redirect("inAtualizacaoCadastral.aspx?ctra=" & Trim(strContratos) & "&id=" & Trim(objDR("ID").ToString) & "&cod=" & UpperTrim(txtCodigo.Text) & "&atz=1", False)
                Server.Transfer("inAtualizacaoCadastral.aspx?ctra=" & Trim(strContratos) & "&id=" & Trim(objDR("ID").ToString) & "&cod=" & UpperTrim(txtCodigo.Text) & "&atz=1", False)

            Else
                'Response.Redirect("inAtualizacaoCadastral.aspx?ctra=" & Trim(strContratos) & "&id=" & Trim(objDR("ID").ToString) & "&cod=" & UpperTrim(txtCodigo.Text) & "&atz=0", False)
                Server.Transfer("inAtualizacaoCadastral.aspx?ctra=" & Trim(strContratos) & "&id=" & Trim(objDR("ID").ToString) & "&cod=" & UpperTrim(txtCodigo.Text) & "&atz=0", False)
            End If

            objDR.Close()
            objDRAC.Close()

        End If

        'Catch ex As Exception
        '    ShowMessage("Sistema Inativo. Por favor tente novamente em alguns minutos.")
        '    SendMailErro("Erro na Negociação On-Line." & Chr(13) & Chr(13) & "Número: " & Err.Number & Chr(13) & "Source: " & Err.Source & Chr(13) & "Descrição: " & Err.Description & Chr(13) & "Help: " & Chr(13) & Err.HelpFile & Chr(13) & Err.HelpContext & Chr(13) & Err.LastDllError)
        'End Try

    End Sub

    Protected Sub ShowMessage(ByVal strMsg As String)

        Dim strScript As String = "<script language=JavaScript>alert('" & strMsg & "');</script>"

        If (Not Page.ClientScript.IsStartupScriptRegistered("clientScript")) Then

            Page.ClientScript.RegisterStartupScript(Me.GetType, "clientScript", strScript)

        End If

    End Sub

    Protected Function SoNumeros(ByVal strDado As String) As String

        Dim x As Short = 0

        Try

            For x = 1 To Len(UpperTrim(strDado))
                If IsNumeric(Mid(UpperTrim(strDado), x, 1)) = False Then
                    strDado = Trim(Replace(UpperTrim(strDado), Mid(UpperTrim(strDado), x, 1), ""))
                End If
            Next x

            SoNumeros = strDado

        Catch ex As Exception

            SoNumeros = "Erro"

        End Try

    End Function

    Protected Function UpperTrim(ByVal strText As String) As String

        Try

            UpperTrim = Trim(UCase(strText))

        Catch ex As Exception

            UpperTrim = ""

        End Try

    End Function

    Protected Sub SendMailErro(ByVal strDescricao As String)

        With eMail
            .From = "x_mail@hargos.com.br"
            .Sender = "Negociação On-Line"
            .ToAddress = "rodrigo.barbieri@hargos.com.br"
            .ToName = "Rodrigo Barbieri - Hargos"
            .CC = "robarbieri@globo.com"
            .CCName = "Rodrigo Barbieri - Globo"
            .IsBodyHTML = False
            .Subject = "Erro na Negociação On-Line."
            .Body = Trim(strDescricao)
        End With
        eMail.Send()

    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "clientScript", "<script language=javascript src=Funcoes/funcoes.js></script>")
        btnCPF.Attributes.Add("onclick", " this.disabled = true; " + ClientScript.GetPostBackEventReference(btnCPF, "") + ";")
        Session("Saldo") = ""
        Session("honor") = ""
        Session("PercentHonor") = ""
        Session("valCorrecao") = ""
        Session("Carteira") = ""
        Session("Principal") = ""
        Session("descNeg") = ""
        Session("SimID") = ""
        Session("CPF") = ""
        Session("Cliente") = ""
        Session("id") = ""
        Session("Mail") = ""
        Session("Nome") = ""
        Session("Contrato") = ""
        Session("ID") = ""
        Session("Nascimento") = ""
        Session("Codigo") = ""
        Session("EntradaMin") = ""
        Session("Parc") = ""
    End Sub

    Protected Sub GravarInput(ByVal strContrato As String, ByVal strMsg As String, ByVal strEmissao As String)

        Dim objFSO As Scripting.FileSystemObject
        Dim objText As Scripting.TextStream = Nothing
        Dim strFile As String = ""

        'strFile = "\\MCLAREN\TRANSFER\RB\Input_Neo\Input_NEG_" & Format(CDate(Now), "yyyMMdd") & ".txt"
        strFile = "D:\Input_Negociacao\Input_NEG_" & Format(CDate(Now), "yyyMMdd") & ".txt"
        objFSO = New Scripting.FileSystemObject
        objText = objFSO.OpenTextFile(strFile, Scripting.IOMode.ForAppending, True)
        objText.WriteLine(Trim(strContrato) & ";" & Trim(strMsg))
        objText.Close()

    End Sub

End Class

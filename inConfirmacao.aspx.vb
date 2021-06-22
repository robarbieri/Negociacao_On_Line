Option Explicit On
Option Strict Off

Imports ConnectTo
Imports System.Data.SqlClient
Imports XMail
Imports System.IO

Partial Class inConfirmacao
    Inherits System.Web.UI.Page

    Protected eMail As New XMail.SendMail
    Protected idLogin As Integer = 0
    Protected simId As Integer = 0
    Protected idCliente As Integer = 0
    Protected strContrato As String = ""
    Protected descNeg As String = ""

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        idLogin = Request.QueryString("id")
        simId = Request.QueryString("simid")
        idCliente = Request.QueryString("idcli")
        strContrato = Trim(Request.QueryString("ctra"))
        descNeg = Trim(Session("descNeg"))
        lblParcela.Text = "<br><br>" & descNeg & "<br><br><br>"
        btnImprimir.Visible = False
        LinkNew.Visible = False
        btnConfirmar.Enabled = True
        btnConfirmar.Visible = True
        btnConfirmar.Focus()

        btnConfirmar.Attributes.Add("onclick", " this.disabled = true; " + " btnVoltar.disabled = true; " + ClientScript.GetPostBackEventReference(btnConfirmar, "") + ";")
        btnVoltar.Attributes.Add("onclick", " this.disabled = true; " + " btnConfirmar.disabled = true; " + ClientScript.GetPostBackEventReference(btnVoltar, "") + ";")

    End Sub

    Protected Sub btnVoltar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnVoltar.Click

        btnVoltar.Enabled = False
        Session("Saldo") = ""
        Response.Redirect("inNegociacao.aspx?ctra=" & strContrato & "&idcliente=" & Trim(Session("Cliente")) & "&id=" & idLogin & "&cod=" & Trim(Session("Codigo")), False)

    End Sub

    Protected Sub btnConfirmar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnConfirmar.Click

        Dim Conn As New Comando
        Dim strSql As String = ""
        Dim objDR As SqlDataReader = Nothing

        Try

            If Left(UCase(Trim(lblParcela.Text)), 5) = "NEGOC" Or Trim(lblParcela.Text) = "" Then
                Exit Try
            End If

            btnConfirmar.Enabled = False
            btnVoltar.Enabled = False

            Conn.Banco = "OSWEB"

            strSql = "OS_S_CADASTRO " & _
                     idCliente & "," & _
                     20 & ",'" & _
                     strContrato & "'"
            objDR = Conn.ExecuteQuery(strSql)

            If objDR.Read Then
                Try
                    strSql = "Delete SituacaoOS Where ID_OS = " & Trim(objDR("ID").ToString)
                    Conn.Execute(strSql)
                    strSql = "OS_D " & Trim(objDR("ID").ToString)
                    Conn.Execute(strSql)
                Catch ex As Exception
                End Try
            End If

            objDR.Close()

            strSql = "OS_IU " & _
                     "0" & "," & _
                     idCliente & "," & _
                     20 & "," & _
                     350 & "," & _
                     394 & "," & _
                     "'" & "0 - Urgente" & "','" & _
                     descNeg & "'"
            Conn.Execute(strSql)

            Conn.Banco = "MIRRORWEB"
            strSql = "prc_INS_Negociacao " & simId
            Conn.Execute(strSql)

            'ShowMessage("Negociação Registrada com sucesso!" & Chr(13) & "Dentro de 48hrs você receberá por e-mail o boleto para pagamento." & Chr(13) & "Qualquer dúvida, entre em contato no telefone SP 11-2171-0301, demais localidades 0800-6000-500 ou envie um e-mail para negociacao@hargos.com.br informando seu CPF.")
            lblParcela.Text = "<br><br><b>Negociação Registrada com SUCESSO!</b><br><br>Dentro de 48hrs você receberá por e-mail o boleto para pagamento.<br>Caso tenha parcelado, você receberá o carnê pelos correios após pagamento da primeira parcela.<br><br>Qualquer dúvida, entre em contato no telefone SP 11-2171-0301, demais localidades 0800-6000-500 ou envie um e-mail para <b>negociacao@hargos.com.br</b> informando seu <b>CPF</b> e o <b>código: " & idLogin & "</b>.<br><br><b>Para sua segurança, imprima esta página.</b><br><br><br>"

            Try

                Dim strFile As String = Server.MapPath("Layout_Mail.html")
                Dim FileStream As StreamReader = File.OpenText(strFile)
                Dim strDados As String = FileStream.ReadToEnd()

                strDados = Replace(strDados, "$NOME_DEVEDOR$", Session("Nome"))
                strDados = Replace(strDados, "$CODIGO$", idLogin)
                strDados = Replace(strDados, "$DADOS_NEG$", descNeg)

                SendMailSucesso(strDados)
                FileStream.Close()

            Catch ex As Exception

                SendMailSucesso("<br><b>Negociação Registrada com SUCESSO!</b><br><br>Dentro de 48hrs você receberá por e-mail o boleto para pagamento.<br>Caso tenha parcelado, você receberá o carnê pelos correios após pagamento da primeira parcela.<br><br>Qualquer dúvida, entre em contato no telefone SP 11-2171-0301, demais localidades 0800-6000-500 ou envie um e-mail para <b>negociacao@hargos.com.br</b> informando seu <b>CPF</b> e o <b>código: " & idLogin & "</b>.<br><br><b>Para sua segurança, imprima esta página.</b><br><br>" & "<b>Dados da Negociação:</b><br><br>" & descNeg)

            End Try

            GravarInput(strContrato, "<b>Negociação On-Line Registrada<br>Descrição:<br>" & descNeg & "<br>Data: " & Format(CDate(Now), "dd/MM/yyy HH:mm:ss") & "</b>", "N_ONLINE")

            btnVoltar.Visible = False
            btnConfirmar.Visible = False
            btnImprimir.Visible = True
            LinkNew.PostBackUrl = "https://dbmirror.hargos.com.br/OnLine/inCPF.aspx"
            LinkNew.Visible = True

        Catch ex As Exception
            lblParcela.Text = "<br><br>Erro na Negociação On-Line.<br><br>Número: " & Err.Number & "<br>Source: " & Err.Source & "<br>Descrição: " & Err.Description & "<br>Help: <br>" & Err.HelpFile & "<br>" & Err.HelpContext & "<br>" & Err.LastDllError
            SendMailErro("Erro na Negociação On-Line." & Chr(13) & Chr(13) & "Número: " & Err.Number & Chr(13) & "Source: " & Err.Source & Chr(13) & "Descrição: " & Err.Description & Chr(13) & "Help: " & Chr(13) & Err.HelpFile & Chr(13) & Err.HelpContext & Chr(13) & Err.LastDllError)
            ShowMessage("Sistema Inativo. Por favor tente novamente em alguns minutos ou entre em contato com nosso atendimento no telefone 11-2171-0301(SP) ou 0800-6000-500(Demais Localidades).")
            'Response.Redirect("inCPF.aspx")
        End Try

    End Sub

    Protected Sub ShowMessage(ByVal strMsg As String)

        Dim strScript As String = "<script language=JavaScript>alert('" & strMsg & "');</script>"

        If (Not Page.ClientScript.IsStartupScriptRegistered("clientScript")) Then

            Page.ClientScript.RegisterStartupScript(Me.GetType, "clientScript", strScript)

        End If

    End Sub

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

    Protected Sub SendMailSucesso(ByVal strDescricao As String)

        Dim strNome As String = ""
        Dim strMail As String = ""
        'Dim arrMail() As String
        Dim strSql As String = ""
        Dim Conn As New Comando
        Dim objDR As SqlDataReader = Nothing
        Dim x As Integer = 0
        'Dim y As Integer = 0
        'Dim blnMultiple As Boolean = False

        Try
            strNome = Trim(Session("Nome"))

            '*'*
            'Session("Mail") = "mara.silva@hargos.com.br"

            If Trim(Session("Mail")) = "" Then

                Conn.Banco = "MIRRORWEB"
                strSql = "Select DEVE_EMAIL Mail From tb_eMail Where DEVE_CGCCPF = '" & Trim(Session("CPF")) & "'"
                objDR = Conn.ExecuteQuery(strSql)

                Do While objDR.Read
                    'If x > 0 Then strMail = strMail & ";"
                    If x > 0 Then Exit Do
                    strMail = strMail & Trim(UCase(objDR("Mail")))
                    x = x + 1
                Loop

                If Trim(strMail) = "" Then
                    SendMailErro("E-Mail não localizado para o CPF " & Trim(Session("CPF")) & ".")
                    Exit Try
                End If
            Else
                strMail = Trim(Session("Mail"))
            End If

            'If InStr(strMail, ";") > 0 Then
            '    arrMail = Split(strMail, ";")
            '    blnMultiple = True
            '    y = UBound(arrMail)
            'End If

            With eMail
                .From = "negociacao@hargos.com.br"
                .HighPriority = True
                .Sender = "Negociação On-Line - Hargos"
                .ToAddress = Trim(strMail)
                .ToName = Trim(strNome)
                .CC = "negociacao@hargos.com.br"
                .CCName = "Negociação On-Line - Hargos"
                .IsBodyHTML = True
                '.Subject = "Negociação On-Line Registrada"
                .Subject = "Confirmação de Negociação On-Line"
                .Body = Trim(strDescricao)
            End With
            eMail.Send()

        Catch ex As Exception
            SendMailErro("Erro na Negociação On-Line." & Chr(13) & Chr(13) & "Número: " & Err.Number & Chr(13) & "Source: " & Err.Source & Chr(13) & "Descrição: " & Err.Description & Chr(13) & "Help: " & Chr(13) & Err.HelpFile & Chr(13) & Err.HelpContext & Chr(13) & Err.LastDllError)
        End Try

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

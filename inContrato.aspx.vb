Option Explicit On
Option Strict Off

Imports ConnectTo
Imports System.Data.SqlClient
Imports XMail

Partial Class inContrato
    Inherits System.Web.UI.Page

    Protected eMail As New XMail.SendMail

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim strContratos As String = ""
        Dim arrContratos() As String
        Dim Mirror As New Comando
        Dim objDR As SqlDataReader = Nothing
        Dim strSql As String = ""
        Dim x As Integer = 0

        Try

            btnContrato.Attributes.Add("onclick", " this.disabled = true; " + ClientScript.GetPostBackEventReference(btnContrato, "") + ";")

            strContratos = Trim(Request.QueryString("ctra"))
            arrContratos = Split(strContratos, ",")

            For x = 0 To UBound(arrContratos)

                cboContratos.Items.Add(Incremente(arrContratos(x), 4, ".", True))

            Next x

            btnContrato.Focus()

        Catch ex As Exception
            SendMailErro("Erro na Negociação On-Line." & Chr(13) & Chr(13) & "Número: " & Err.Number & Chr(13) & "Source: " & Err.Source & Chr(13) & "Descrição: " & Err.Description & Chr(13) & "Help: " & Chr(13) & Err.HelpFile & Chr(13) & Err.HelpContext & Chr(13) & Err.LastDllError)
            ShowMessage("Sistema Inativo. Por favor tente novamente em alguns minutos.")
            'Response.Redirect("inCPF.aspx")
        End Try

    End Sub

    Protected Sub btnContrato_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnContrato.Click

        Dim Mirror As New Comando
        Dim strSql As String = ""
        Dim objDR As SqlDataReader = Nothing
        Dim objDRAC As SqlDataReader = Nothing

        'Try

        btnContrato.Enabled = False

        Mirror.Banco = "MIRRORWEB"
        strSql = "prc_INS_NEGLOGIN '" & SoNumeros(Session("CPF")) & "','" & SoNumeros(Trim(cboContratos.SelectedItem.Value.ToString)) & "'"
        Mirror.Execute(strSql)

        strSql = "prc_SEL_NEGLOGIN '" & SoNumeros(Trim(cboContratos.SelectedItem.Value.ToString)) & "'"
        objDR = Mirror.ExecuteQuery(strSql)

        If Not objDR.Read Then
            SendMailErro("Erro na Negociação On-Line." & Chr(13) & Chr(13) & "Número: " & Err.Number & Chr(13) & "Source: " & Err.Source & Chr(13) & "Descrição: " & Err.Description & Chr(13) & "Help: " & Chr(13) & Err.HelpFile & Chr(13) & Err.HelpContext & Chr(13) & Err.LastDllError)
            ShowMessage("Sistema Inativo. Por favor tente novamente em alguns minutos.")
            'Response.Redirect("inCPF.aspx")
            'Exit Try
            btnContrato.Enabled = True
            Exit Sub
        End If

        Session("Contrato") = Trim(cboContratos.SelectedItem.Value.ToString)
        Session("ID") = Trim(objDR("ID").ToString)

        GravarInput(Trim(cboContratos.SelectedItem.Value.ToString), "<b>Acesso à Negociação On-Line.<br>ID: " & Trim(objDR("ID").ToString) & "<br>Data: " & Format(CDate(Now), "dd/MM/yyy HH:mm:ss") & "</b>", "N_ONLINE")

        strSql = "prc_SEL_AtualizacaoCadastral " & SoNumeros(Session("CPF"))
        objDRAC = Mirror.ExecuteQuery(strSql)

        If Not objDRAC.Read Then
            Response.Redirect("inAtualizacaoCadastral.aspx?ctra=" & Trim(Replace(cboContratos.SelectedItem.Value.ToString, ".", "")) & "&id=" & Trim(objDR("ID").ToString) & "&cod=" & Trim(Request.QueryString("cod")) & "&atz=1", False)
        Else
            Response.Redirect("inAtualizacaoCadastral.aspx?ctra=" & Trim(Replace(cboContratos.SelectedItem.Value.ToString, ".", "")) & "&id=" & Trim(objDR("ID").ToString) & "&cod=" & Trim(Request.QueryString("cod")) & "&atz=0", False)
        End If

        'Catch ex As Exception
        '    SendMailErro("Erro na Negociação On-Line." & Chr(13) & Chr(13) & "Número: " & Err.Number & Chr(13) & "Source: " & Err.Source & Chr(13) & "Descrição: " & Err.Description & Chr(13) & "Help: " & Chr(13) & Err.HelpFile & Chr(13) & Err.HelpContext & Chr(13) & Err.LastDllError)
        '    ShowMessage("Sistema Inativo. Por favor tente novamente em alguns minutos.")
        '    Response.Redirect("inCPF.aspx")
        'End Try

    End Sub

    Protected Sub ShowMessage(ByVal strMsg As String)

        Dim strScript As String = "<script language=JavaScript>alert('" & strMsg & "');</script>"

        If (Not Page.ClientScript.IsStartupScriptRegistered("clientScript")) Then

            Page.ClientScript.RegisterStartupScript(Me.GetType, "clientScript", strScript)

        End If

    End Sub

    Protected Function Incremente(ByRef strTexto As Object, ByRef intCada As Short, ByRef strIncremente As String, ByRef blnRemoverUltimo As Boolean) As Object

        Dim x As Short
        Dim strResult As String = ""

        If InStr(strTexto, strIncremente) = 0 Then

            strTexto = Trim(strTexto)

            For x = 1 To Len(strTexto)

                strResult = strResult & Mid(strTexto, x, 1)

                If InStr(CStr(x / intCada), ",") = 0 Then

                    If x = Len(strTexto) Then
                        If blnRemoverUltimo = False Then
                            strResult = strResult & "."
                        End If
                    Else
                        strResult = strResult & "."
                    End If

                End If

            Next x

            Incremente = Trim(strResult)

        Else

            Incremente = strTexto

        End If

    End Function

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

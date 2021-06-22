Option Explicit On
Option Strict Off

Imports ConnectTo
Imports System.Data.SqlClient

Partial Class inAtualizacaoCadastral
    Inherits System.Web.UI.Page

    Protected idLogin As Long = 0
    Protected eMail As New XMail.SendMail

    Protected Sub btnEndereco_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEndereco.Click

        Dim strContrato As String = ""
        Dim idDeve As Long = 0
        Dim Conn As New Comando
        Dim objDR As SqlDataReader = Nothing
        Dim strSql As String = ""
        Dim x As Integer = 0
        'Dim blnPreenchido As Boolean = True
        Dim intResponsavel As Integer = 0
        Dim strAssunto As String = ""
        Dim strDescricao As String = ""
        Dim Funcoes As New Funcoes.Local
        Dim strDadosInput As String = ""

        'Try

        strContrato = Trim(Request.QueryString("ctra"))
        idLogin = Trim(Session("id"))

        If (Trim(txtEMail.Text) = "" Or _
           InStr(txtEMail.Text, "@") = 0 Or _
           InStr(UCase(txtEMail.Text), ".COM") = 0) Then
            ShowMessage("Preencha corretamente o campo E-Mail.")
            txtEMail.Focus()
            Exit Sub
        End If

        If Trim(txtEndResLogr.Text) = "" Then
            ShowMessage("Preencha corretamente o campo Endereço Residencial - Logradouro.")
            txtEndResLogr.Focus()
            Exit Sub
        End If

        If Trim(txtEndResNum.Text) = "" Or _
            Not IsNumeric(Trim(txtEndResNum.Text)) Then
            ShowMessage("Preencha corretamente o campo Endereço Residencial - Número.")
            txtEndResNum.Focus()
            Exit Sub
        End If

        If Trim(txtEndResBairro.Text) = "" Then
            ShowMessage("Preencha corretamente o campo Bairro Residencial.")
            txtEndResBairro.Focus()
            Exit Sub
        End If

        If Trim(txtEndResCEP.Text) = "" Or _
            Len(Funcoes.SoNumeros(Trim(txtEndResCEP.Text))) <> 8 Then
            ShowMessage("Preencha corretamente o campo CEP Residencial. Deve conter 8 caracteres numéricos.")
            txtEndResCEP.Focus()
            Exit Sub
        End If

        If Trim(txtEndResCidade.Text) = "" Then
            ShowMessage("Preencha corretamente o campo Cidade Residencial.")
            txtEndResCidade.Focus()
            Exit Sub
        End If

        If (Trim(txtDDDRes1.Text) = "" Or _
            Len(Funcoes.SoNumeros(Trim(txtDDDRes1.Text))) <> 2) Then
            ShowMessage("Preencha corretamente o campo DDD Residencial. Deve conter 2 caracteres numéricos.")
            txtDDDRes1.Focus()
            Exit Sub
        End If

        If Trim(txtFoneRes1.Text) = "" Or _
           Len(Funcoes.SoNumeros(Trim(txtFoneRes1.Text))) <> 8 Then
            ShowMessage("Preencha corretamente o campo Fone Residencial. Deve conter 8 caracteres numéricos.")
            txtFoneRes1.Focus()
            Exit Sub
        End If

        'If blnPreenchido = False Then
        '    ShowMessage("Preencha corretamente os campos marcados com *.")
        '    'Exit Try
        '    Exit Sub
        'End If

        Session("Mail") = Trim(txtEMail.Text)

        Conn.Banco = "NEOWEB"
        'Conn.Banco = "NEOWEB_REAL"
        strSql = "Select TOP 1 DEVE_Id From Contratos Where CTRA_Numero = '" & strContrato & "'"
        objDR = Conn.ExecuteQuery(strSql)
        objDR.Read()

        idDeve = CLng(Trim(objDR(0)))

        Conn.Banco = "MIRRORWEB"

        txtEndResLogr.Text = Replace(txtEndResLogr.Text, ".", " ")
        txtEndResLogr.Text = Replace(txtEndResLogr.Text, "  ", " ")
        txtEndComLogr.Text = Replace(txtEndComLogr.Text, ".", " ")
        txtEndComLogr.Text = Replace(txtEndComLogr.Text, "  ", " ")

        strDescricao = ""
        strDescricao = "<b>Contrato: </b>" & strContrato & "<br>"
        strDescricao = strDescricao & "<b>E-Mail: </b>" & UpperTrim(txtEMail.Text) & "<br>"

        If Trim(txtEMail.Text) <> "" Then
            Try
                strSql = "Insert Into tb_eMail Values(" & _
                         "Right('00000000000' + '" & Session("CPF") & "',11),'" & _
                         UpperTrim(txtEMail.Text) & "')"
                '*'*'Conn.Execute(strSql)
                Conn.Execute(strSql)
            Catch ex As Exception
            End Try
        End If

        If Trim(txtEndResLogr.Text) <> "" Then

            Try
                strSql = "prc_INS_Endereco 1,'" & _
                         UpperTrim(txtEndResLogr.Text) & "','" & _
                         Trim(Funcoes.SoNumeros(txtEndResNum.Text)) & "','" & _
                         UpperTrim(txtEndResCompl.Text) & "','" & _
                         UpperTrim(txtEndResBairro.Text) & "','" & _
                         Trim(Funcoes.SoNumeros(txtEndResCEP.Text)) & "','" & _
                         UpperTrim(txtEndResCidade.Text) & "','" & _
                         UpperTrim(cboResUF.SelectedItem.Value) & "'," & _
                         idDeve
                Conn.Execute(strSql)
            Catch ex As Exception
            End Try

            strDescricao = strDescricao & "<b>Endereço Residencial:</b> " & UpperTrim(txtEndResLogr.Text) & "<br>"
            strDescricao = strDescricao & "<b>Número: </b>" & Trim(Funcoes.SoNumeros(txtEndResNum.Text)) & "<br>"
            If Trim(txtEndResCompl.Text) <> "" Then strDescricao = strDescricao & "<b>Complemento: </b>" & UpperTrim(txtEndResCompl.Text) & "<br>"
            strDescricao = strDescricao & "<b>Bairro: </b>" & UpperTrim(txtEndResBairro.Text) & "<br>"
            strDescricao = strDescricao & "<b>CEP: </b>" & Trim(Funcoes.SoNumeros(txtEndResCEP.Text)) & "<br>"
            strDescricao = strDescricao & "<b>Cidade: </b>" & UpperTrim(txtEndResCidade.Text) & "<br>"
            strDescricao = strDescricao & "<b>UF: </b>" & UpperTrim(cboResUF.SelectedItem.Value)

        End If

        If Trim(txtEndComLogr.Text) <> "" Then

            Try
                strSql = "prc_INS_Endereco 2,'" & _
                         UpperTrim(txtEndComLogr.Text) & "','" & _
                         Trim(Funcoes.SoNumeros(txtEndComNum.Text)) & "','" & _
                         UpperTrim(txtEndComCompl.Text) & "','" & _
                         UpperTrim(txtEndComBairro.Text) & "','" & _
                         Trim(Funcoes.SoNumeros(txtEndComCep.Text)) & "','" & _
                         UpperTrim(txtEndComCidade.Text) & "','" & _
                         UpperTrim(cboComUF.SelectedItem.Value) & "'," & _
                         idDeve
                Conn.Execute(strSql)
            Catch ex As Exception
            End Try

            strDescricao = strDescricao & "<br><br>"
            strDescricao = strDescricao & "<b>Endereço Comercial:</b> " & UpperTrim(txtEndComLogr.Text) & "<br>"
            strDescricao = strDescricao & "<b>Número: </b>" & Trim(Funcoes.SoNumeros(txtEndComNum.Text)) & "<br>"
            If Trim(txtEndComCompl.Text) <> "" Then strDescricao = strDescricao & "<b>Complemento: </b>" & UpperTrim(txtEndComCompl.Text) & "<br>"
            strDescricao = strDescricao & "<b>Bairro: </b>" & UpperTrim(txtEndComBairro.Text) & "<br>"
            strDescricao = strDescricao & "<b>CEP: </b>" & Trim(Funcoes.SoNumeros(txtEndComCep.Text)) & "<br>"
            strDescricao = strDescricao & "<b>Cidade: </b>" & UpperTrim(txtEndComCidade.Text) & "<br>"
            strDescricao = strDescricao & "<b>UF: </b>" & UpperTrim(cboComUF.SelectedItem.Value)

        End If

        Try

            If Len(SoNumeros(txtDDDRes1.Text)) = 2 Then
                Try
                    strSql = "prc_INS_Telefone 1," & idDeve & ",'" & _
                                             Trim(Funcoes.SoNumeros(txtDDDRes1.Text)) & "','" & _
                                             Trim(Funcoes.SoNumeros(txtFoneRes1.Text)) & "'"
                    Conn.Execute(strSql)
                Catch ex As Exception
                End Try

                strDescricao = strDescricao & "<br><br>"
                strDescricao = strDescricao & "<b>DDD Residencial:</b> " & Trim(Funcoes.SoNumeros(txtDDDRes1.Text)) & "<br>"
                strDescricao = strDescricao & "<b>Telefone Residencial: </b>" & Trim(Funcoes.SoNumeros(txtFoneRes1.Text)) & "<br>"
            End If

        Catch ex As Exception
        End Try

        Try

            If Len(SoNumeros(txtDDDRes2.Text)) = 2 Then
                Try
                    strSql = "prc_INS_Telefone 4," & idDeve & ",'" & _
                                             Trim(Funcoes.SoNumeros(txtDDDRes2.Text)) & "','" & _
                                             Trim(Funcoes.SoNumeros(txtFoneRes2.Text)) & "'"
                    Conn.Execute(strSql)
                Catch ex As Exception
                End Try

                strDescricao = strDescricao & "<br>"
                strDescricao = strDescricao & "<b>DDD Residencial:</b> " & Trim(Funcoes.SoNumeros(txtDDDRes2.Text)) & "<br>"
                strDescricao = strDescricao & "<b>Fax Residencial: </b>" & Trim(Funcoes.SoNumeros(txtFoneRes2.Text)) & "<br>"
            End If

        Catch ex As Exception
        End Try

        Try

            If Len(SoNumeros(txtDDDCel1.Text)) = 2 Then
                Try
                    strSql = "prc_INS_Telefone 2," & idDeve & ",'" & _
                                             Trim(Funcoes.SoNumeros(txtDDDCel1.Text)) & "','" & _
                                             Trim(Funcoes.SoNumeros(txtFoneCel1.Text)) & "'"
                    Conn.Execute(strSql)
                Catch ex As Exception
                End Try
                strDescricao = strDescricao & "<br>"
                strDescricao = strDescricao & "<b>DDD Celular 1:</b> " & Trim(Funcoes.SoNumeros(txtDDDCel1.Text)) & "<br>"
                strDescricao = strDescricao & "<b>Telefone Celular 1: </b>" & Trim(Funcoes.SoNumeros(txtFoneCel1.Text)) & "<br>"
            End If

        Catch ex As Exception
        End Try

        Try

            If Len(SoNumeros(txtDDDCel2.Text)) = 2 Then
                Try
                    strSql = "prc_INS_Telefone 2," & idDeve & ",'" & _
                                             Trim(Funcoes.SoNumeros(txtDDDCel2.Text)) & "','" & _
                                             Trim(Funcoes.SoNumeros(txtFoneCel2.Text)) & "'"
                    Conn.Execute(strSql)
                Catch ex As Exception
                End Try
                strDescricao = strDescricao & "<br><br>"
                strDescricao = strDescricao & "<b>DDD Celular 2:</b> " & Trim(Funcoes.SoNumeros(txtDDDCel2.Text)) & "<br>"
                strDescricao = strDescricao & "<b>Telefone Celular 2: </b>" & Trim(Funcoes.SoNumeros(txtFoneCel2.Text)) & "<br>"
            End If

        Catch ex As Exception
        End Try

        Try

            If Len(SoNumeros(txtDDDCom1.Text)) = 2 Then
                Try
                    strSql = "prc_INS_Telefone 3," & idDeve & ",'" & _
                                             Trim(Funcoes.SoNumeros(txtDDDCom1.Text)) & "','" & _
                                             Trim(Funcoes.SoNumeros(txtFoneCom1.Text)) & "'"
                    Conn.Execute(strSql)
                Catch ex As Exception
                End Try
                strDescricao = strDescricao & "<br>"
                strDescricao = strDescricao & "<b>DDD Comercial:</b> " & Trim(Funcoes.SoNumeros(txtDDDCom1.Text)) & "<br>"
                strDescricao = strDescricao & "<b>Telefone Comercial: </b>" & Trim(Funcoes.SoNumeros(txtFoneCom1.Text)) & "<br>"
            End If

        Catch ex As Exception
        End Try

        Try

            If Len(SoNumeros(txtDDDCom2.Text)) = 2 Then
                Try
                    strSql = "prc_INS_Telefone 5," & idDeve & ",'" & _
                                             Trim(Funcoes.SoNumeros(txtDDDCom2.Text)) & "','" & _
                                             Trim(Funcoes.SoNumeros(txtFoneCom2.Text)) & "'"
                    Conn.Execute(strSql)
                Catch ex As Exception
                End Try
                strDescricao = strDescricao & "<br>"
                strDescricao = strDescricao & "<b>DDD Comercial:</b> " & Trim(Funcoes.SoNumeros(txtDDDCom2.Text)) & "<br>"
                strDescricao = strDescricao & "<b>Fax Comercial: </b>" & Trim(Funcoes.SoNumeros(txtFoneCom2.Text)) & "<br>"
            End If

        Catch ex As Exception
        End Try

        Try
            Conn.Banco = "NEOWEB"
            'Conn.Banco = "NEOWEB_REAL"
            strSql = "Select TOP 1 DEVE_Nome,DEVE_CGCCPF From Devedores Where DEVE_Id = " & idDeve
            objDR = Conn.ExecuteQuery(strSql)

            If objDR.Read Then

                Dim intPos As Integer = 0
                Dim strTipo As String = ""
                Dim strEnd As String = ""

                intPos = InStr(Trim(txtEndResLogr.Text), " ")
                strTipo = UpperTrim(Left(Trim(txtEndResLogr.Text), intPos - 1))
                strEnd = UpperTrim(Mid(Trim(txtEndResLogr.Text), intPos + 1, Len(Trim(txtEndResLogr.Text))))
                Try
                    strSql = "prc_INS_Pesquisa '" & SoNumeros(txtDDDRes1.Text) & "','" & _
                                                    SoNumeros(txtFoneRes1.Text) & "','" & _
                                                    UpperTrim(objDR("DEVE_Nome")) & "','" & _
                                                    strTipo & "','" & _
                                                    strEnd & "','" & _
                                                    SoNumeros(txtEndResNum.Text) & "','" & _
                                                    UpperTrim(txtEndResCompl.Text) & "','" & _
                                                    UpperTrim(txtEndResBairro.Text) & "','" & _
                                                    UpperTrim(txtEndResCidade.Text) & "','" & _
                                                    UpperTrim(cboResUF.SelectedItem.Value) & "','" & _
                                                    SoNumeros(txtEndResCEP.Text) & "','" & _
                                                    SoNumeros(objDR("DEVE_CGCCPF")) & "'," & _
                                                    10

                    '*'*'Conn.Execute(strSql)
                    Conn.Execute(strSql)
                Catch ex As Exception
                End Try
                If Len(SoNumeros(txtDDDRes2.Text)) = 2 And Len(SoNumeros(txtFoneRes2.Text)) = 8 Then
                    Try
                        strSql = "prc_INS_Pesquisa '" & SoNumeros(txtDDDRes2.Text) & "','" & _
                                                        SoNumeros(txtFoneRes2.Text) & "','" & _
                                                        UpperTrim(objDR("DEVE_Nome")) & "','" & _
                                                        strTipo & "','" & _
                                                        strEnd & "','" & _
                                                        SoNumeros(txtEndResNum.Text) & "','" & _
                                                        UpperTrim(txtEndResCompl.Text) & "','" & _
                                                        UpperTrim(txtEndResBairro.Text) & "','" & _
                                                        UpperTrim(txtEndResCidade.Text) & "','" & _
                                                        UpperTrim(cboResUF.SelectedItem.Value) & "','" & _
                                                        SoNumeros(txtEndResCEP.Text) & "','" & _
                                                        SoNumeros(objDR("DEVE_CGCCPF")) & "'," & _
                                                        10

                        '*'*'Conn.Execute(strSql)
                        Conn.Execute(strSql)
                    Catch ex As Exception
                    End Try
                End If

                If Len(SoNumeros(txtDDDCel1.Text)) = 2 And Len(SoNumeros(txtFoneCel1.Text)) = 8 Then
                    Try
                        strSql = "prc_INS_Pesquisa '" & SoNumeros(txtDDDCel1.Text) & "','" & _
                                                SoNumeros(txtFoneCel1.Text) & "','" & _
                                                UpperTrim(objDR("DEVE_Nome")) & "','" & _
                                                strTipo & "','" & _
                                                strEnd & "','" & _
                                                SoNumeros(txtEndResNum.Text) & "','" & _
                                                UpperTrim(txtEndResCompl.Text) & "','" & _
                                                UpperTrim(txtEndResBairro.Text) & "','" & _
                                                UpperTrim(txtEndResCidade.Text) & "','" & _
                                                UpperTrim(cboResUF.SelectedItem.Value) & "','" & _
                                                SoNumeros(txtEndResCEP.Text) & "','" & _
                                                SoNumeros(objDR("DEVE_CGCCPF")) & "'," & _
                                                10

                        '*'*'Conn.Execute(strSql)
                        Conn.Execute(strSql)
                    Catch ex As Exception
                    End Try
                End If

                If Len(SoNumeros(txtDDDCel2.Text)) = 2 And Len(SoNumeros(txtFoneCel2.Text)) = 8 Then
                    Try
                        strSql = "prc_INS_Pesquisa '" & SoNumeros(txtDDDCel2.Text) & "','" & _
                                                SoNumeros(txtFoneCel2.Text) & "','" & _
                                                UpperTrim(objDR("DEVE_Nome")) & "','" & _
                                                strTipo & "','" & _
                                                strEnd & "','" & _
                                                SoNumeros(txtEndResNum.Text) & "','" & _
                                                UpperTrim(txtEndResCompl.Text) & "','" & _
                                                UpperTrim(txtEndResBairro.Text) & "','" & _
                                                UpperTrim(txtEndResCidade.Text) & "','" & _
                                                UpperTrim(cboResUF.SelectedItem.Value) & "','" & _
                                                SoNumeros(txtEndResCEP.Text) & "','" & _
                                                SoNumeros(objDR("DEVE_CGCCPF")) & "'," & _
                                                10

                        '*'*'Conn.Execute(strSql)
                        Conn.Execute(strSql)
                    Catch ex As Exception
                    End Try
                End If

                If Len(SoNumeros(txtDDDCom1.Text)) = 2 And Len(SoNumeros(txtFoneCom1.Text)) = 8 Then
                    Try
                        strSql = "prc_INS_Pesquisa '" & SoNumeros(txtDDDCom1.Text) & "','" & _
                                                SoNumeros(txtFoneCom1.Text) & "','" & _
                                                UpperTrim(objDR("DEVE_Nome")) & "','"
                        If Trim(txtEndComLogr.Text) <> "" Then
                            intPos = InStr(Trim(txtEndComLogr.Text), " ")
                            strTipo = UpperTrim(Left(Trim(txtEndComLogr.Text), intPos - 1))
                            strEnd = UpperTrim(Mid(Trim(txtEndComLogr.Text), intPos + 1, Len(Trim(txtEndComLogr.Text))))
                            strSql = strSql & strTipo & "','" & _
                                                        strEnd & "','" & _
                                                        SoNumeros(txtEndComNum.Text) & "','" & _
                                                        UpperTrim(txtEndComCompl.Text) & "','" & _
                                                        UpperTrim(txtEndComBairro.Text) & "','" & _
                                                        UpperTrim(txtEndComCidade.Text) & "','" & _
                                                        UpperTrim(cboComUF.SelectedItem.Value) & "','" & _
                                                        SoNumeros(txtEndComCep.Text) & "','" & _
                                                        SoNumeros(objDR("DEVE_CGCCPF")) & "'," & _
                                                        10
                        Else
                            strSql = strSql & strTipo & "','" & _
                                                        strEnd & "','" & _
                                                        SoNumeros(txtEndResNum.Text) & "','" & _
                                                        UpperTrim(txtEndResCompl.Text) & "','" & _
                                                        UpperTrim(txtEndResBairro.Text) & "','" & _
                                                        UpperTrim(txtEndResCidade.Text) & "','" & _
                                                        UpperTrim(cboResUF.SelectedItem.Value) & "','" & _
                                                        SoNumeros(txtEndResCEP.Text) & "','" & _
                                                        SoNumeros(objDR("DEVE_CGCCPF")) & "'," & _
                                                        10
                        End If
                        '*'*'Conn.Execute(strSql)
                        Conn.Execute(strSql)
                    Catch ex As Exception
                    End Try
                End If

                If Len(SoNumeros(txtDDDCom2.Text)) = 2 And Len(SoNumeros(txtFoneCom2.Text)) = 8 Then
                    Try
                        strSql = "prc_INS_Pesquisa '" & SoNumeros(txtDDDCom2.Text) & "','" & _
                                                SoNumeros(txtFoneCom2.Text) & "','" & _
                                                UpperTrim(objDR("DEVE_Nome")) & "','"
                        If Trim(txtEndComLogr.Text) <> "" Then
                            intPos = InStr(Trim(txtEndComLogr.Text), " ")
                            strTipo = UpperTrim(Left(Trim(txtEndComLogr.Text), intPos - 1))
                            strEnd = UpperTrim(Mid(Trim(txtEndComLogr.Text), intPos + 1, Len(Trim(txtEndComLogr.Text))))
                            strSql = strSql & strTipo & "','" & _
                                                        strEnd & "','" & _
                                                        SoNumeros(txtEndComNum.Text) & "','" & _
                                                        UpperTrim(txtEndComCompl.Text) & "','" & _
                                                        UpperTrim(txtEndComBairro.Text) & "','" & _
                                                        UpperTrim(txtEndComCidade.Text) & "','" & _
                                                        UpperTrim(cboComUF.SelectedItem.Value) & "','" & _
                                                        SoNumeros(txtEndComCep.Text) & "','" & _
                                                        SoNumeros(objDR("DEVE_CGCCPF")) & "'," & _
                                                        10
                        Else
                            strSql = strSql & strTipo & "','" & _
                                                        strEnd & "','" & _
                                                        SoNumeros(txtEndResNum.Text) & "','" & _
                                                        UpperTrim(txtEndResCompl.Text) & "','" & _
                                                        UpperTrim(txtEndResBairro.Text) & "','" & _
                                                        UpperTrim(txtEndResCidade.Text) & "','" & _
                                                        UpperTrim(cboResUF.SelectedItem.Value) & "','" & _
                                                        SoNumeros(txtEndResCEP.Text) & "','" & _
                                                        SoNumeros(objDR("DEVE_CGCCPF")) & "'," & _
                                                        10
                        End If
                        '*'*'Conn.Execute(strSql)
                        Conn.Execute(strSql)
                    Catch ex As Exception
                    End Try
                End If

            End If

        Catch ex As Exception
        End Try

        intResponsavel = 394
        'intResponsavel = 350
        strAssunto = "0 - Urgente"

        'Conn.Banco = "NEOWEB"
        'strSql = "Select C.CONT_Id, B.CART_Id From Contratos A (NOLOCK) Join " & _
        '         "Carteiras B (NOLOCK) On B.CART_Id = A.CART_Id Join " & _
        '         "Contratante C (NOLOCK) On C.CONT_Id = B.CONT_Id " & _
        '         "Where A.CTRA_Numero = '" & strContrato & "'"
        'objDR = Conn.ExecuteQuery(strSql)
        'objDR.Read()

        'Session("Carteira") = Trim(objDR("CART_Id"))

        Conn.Banco = "OSWEB"

        'strSql = "Select ID_Cliente From Clientes Where Cod_Cliente = " & objDR("CONT_Id")
        'objDR = Conn.ExecuteQuery(strSql)
        'objDR.Read()

        'Session("Cliente") = Trim(objDR("ID_Cliente"))

        strSql = "OS_S_CADASTRO " & _
                 Session("Cliente") & "," & _
                 19 & ",'" & _
                 strContrato & "'"
        objDR = Conn.ExecuteQuery(strSql)

        If Not objDR.Read Then
            Try
                strSql = "OS_IU " & _
                         "0" & "," & _
                         Session("Cliente") & "," & _
                         19 & "," & _
                         350 & "," & _
                         intResponsavel & "," & _
                         "'" & strAssunto & "','" & _
                         strDescricao & "'"
                '*'*'Conn.Execute(strSql)
                Conn.Execute(strSql)
            Catch ex As Exception
            End Try
        End If

        'Conn.Banco = "MIRRORWEB"
        'strDadosInput = ""

        'strDadosInput = idLogin & ",'" & _
        '                UpperTrim(txtEMail.Text) & "','" & _
        '                UpperTrim(txtEndResLogr.Text) & "','" & _
        '                SoNumeros(txtEndResNum.Text) & "','" & _
        '                UpperTrim(txtEndResCompl.Text) & "','" & _
        '                UpperTrim(txtEndResBairro.Text) & "','" & _
        '                SoNumeros(txtEndResCEP.Text) & "','" & _
        '                UpperTrim(txtEndResCidade.Text) & "','" & _
        '                UpperTrim(cboResUF.SelectedItem.Value) & "','" & _
        '                UpperTrim(txtEndComLogr.Text) & "','" & _
        '                SoNumeros(txtEndComNum.Text) & "','" & _
        '                UpperTrim(txtEndComCompl.Text) & "','" & _
        '                UpperTrim(txtEndComBairro.Text) & "','" & _
        '                SoNumeros(txtEndComCep.Text) & "','" & _
        '                UpperTrim(txtEndComCidade.Text) & "','" & _
        '                UpperTrim(cboComUF.SelectedItem.Value) & "','" & _
        '                SoNumeros(txtDDDRes1.Text) & "','" & _
        '                SoNumeros(txtFoneRes1.Text) & "','" & _
        '                SoNumeros(txtDDDRes2.Text) & "','" & _
        '                SoNumeros(txtFoneRes2.Text) & "','" & _
        '                SoNumeros(txtDDDCom1.Text) & "','" & _
        '                SoNumeros(txtFoneCom1.Text) & "','" & _
        '                SoNumeros(txtDDDCom2.Text) & "','" & _
        '                SoNumeros(txtFoneCom2.Text) & "','" & _
        '                SoNumeros(txtDDDCel1.Text) & "','" & _
        '                SoNumeros(txtFoneCel1.Text) & "','" & _
        '                SoNumeros(txtDDDCel2.Text) & "','" & _
        '                SoNumeros(txtFoneCel2.Text) & "'"

        'strSql = "prc_INS_NEGATZCADASTRAL " & strDadosInput
        'Conn.Execute(strSql)

        Conn.Banco = "MIRRORWEB"
        strSql = "prc_INS_NegAC " & SoNumeros(Session("CPF"))
        Conn.Execute(strSql)

        GravarInput(strContrato, "<b>Atualização Cadastral pela Negociação On-Line<br>Data: " & Format(CDate(Now), "dd/MM/yyy HH:mm:ss") & "</b>", "N_ONLINE")

        Response.Redirect("inNegociacao.aspx?ctra=" & strContrato & "&idcliente=" & Trim(Session("Cliente")) & "&id=" & idLogin & "&cod=" & Trim(Request.QueryString("cod")), False)

        'Catch ex As Exception
        '    ShowMessage("Sistema Inativo. Por favor tente novamente em alguns minutos.")
        '    Response.Redirect("inCPF.aspx", False)
        'End Try

    End Sub

    Protected Sub ShowMessage(ByVal strMsg As String)

        Dim strScript As String = "<script language=JavaScript>alert('" & strMsg & "');</script>"

        If (Not Page.ClientScript.IsStartupScriptRegistered("clientScript")) Then

            Page.ClientScript.RegisterStartupScript(Me.GetType, "clientScript", strScript)

        End If

    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Dim intAtualizar As Short = 1
        Dim Conn As New Comando
        Dim strSql As String = ""
        Dim objDR As SqlDataReader = Nothing
        Dim strContrato As String = ""
        'Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "clientScript", "<script language=javascript src=Funcoes/funcoes.js></script>")

        btnEndereco.Attributes.Add("onclick", " this.disabled = true; " + ClientScript.GetPostBackEventReference(btnEndereco, "") + ";")

        If Trim(Session("Cliente")) = "" Then

            strContrato = Trim(Request.QueryString("ctra"))
            idLogin = Trim(Session("id"))
            intAtualizar = Trim(Request.QueryString("atz"))

            Conn.Banco = "NEOWEB"
            strSql = "Select C.CONT_Id, B.CART_Id From Contratos A (NOLOCK) Join " & _
                     "Carteiras B (NOLOCK) On B.CART_Id = A.CART_Id Join " & _
                     "Contratante C (NOLOCK) On C.CONT_Id = B.CONT_Id " & _
                     "Where A.CTRA_Numero = '" & strContrato & "'"
            objDR = Conn.ExecuteQuery(strSql)
            objDR.Read()

            Session("Carteira") = Trim(objDR("CART_Id"))

            Conn.Banco = "OSWEB"

            strSql = "Select ID_Cliente From Clientes Where Cod_Cliente = " & objDR("CONT_Id")
            objDR = Conn.ExecuteQuery(strSql)
            objDR.Read()

            Session("Cliente") = Trim(objDR("ID_Cliente"))

        End If

        If intAtualizar = 0 Then
            Response.Redirect("inNegociacao.aspx?ctra=" & strContrato & "&idcliente=" & Trim(Session("Cliente")) & "&id=" & idLogin & "&cod=" & Trim(Request.QueryString("cod")), False)
        Else

            If cboResUF.Items.Count < 29 Then
                cboResUF.Items.Clear()
                cboComUF.Items.Clear()
                cboResUF.Items.Add("AC")
                cboResUF.Items.Add("AL")
                cboResUF.Items.Add("AP")
                cboResUF.Items.Add("AM")
                cboResUF.Items.Add("BA")
                cboResUF.Items.Add("CE")
                cboResUF.Items.Add("DF")
                cboResUF.Items.Add("ES")
                cboResUF.Items.Add("GO")
                cboResUF.Items.Add("MA")
                cboResUF.Items.Add("MT")
                cboResUF.Items.Add("MS")
                cboResUF.Items.Add("MZ")
                cboResUF.Items.Add("MG")
                cboResUF.Items.Add("NI")
                cboResUF.Items.Add("PA")
                cboResUF.Items.Add("PB")
                cboResUF.Items.Add("PR")
                cboResUF.Items.Add("PE")
                cboResUF.Items.Add("PI")
                cboResUF.Items.Add("RJ")
                cboResUF.Items.Add("RN")
                cboResUF.Items.Add("RS")
                cboResUF.Items.Add("RO")
                cboResUF.Items.Add("RR")
                cboResUF.Items.Add("SC")
                cboResUF.Items.Add("SP")
                cboResUF.Items.Add("SE")
                cboResUF.Items.Add("TO")
                cboComUF.Items.Add("AC")
                cboComUF.Items.Add("AL")
                cboComUF.Items.Add("AP")
                cboComUF.Items.Add("AM")
                cboComUF.Items.Add("BA")
                cboComUF.Items.Add("CE")
                cboComUF.Items.Add("DF")
                cboComUF.Items.Add("ES")
                cboComUF.Items.Add("GO")
                cboComUF.Items.Add("MA")
                cboComUF.Items.Add("MT")
                cboComUF.Items.Add("MS")
                cboComUF.Items.Add("MZ")
                cboComUF.Items.Add("MG")
                cboComUF.Items.Add("NI")
                cboComUF.Items.Add("PA")
                cboComUF.Items.Add("PB")
                cboComUF.Items.Add("PR")
                cboComUF.Items.Add("PE")
                cboComUF.Items.Add("PI")
                cboComUF.Items.Add("RJ")
                cboComUF.Items.Add("RN")
                cboComUF.Items.Add("RS")
                cboComUF.Items.Add("RO")
                cboComUF.Items.Add("RR")
                cboComUF.Items.Add("SC")
                cboComUF.Items.Add("SP")
                cboComUF.Items.Add("SE")
                cboComUF.Items.Add("TO")

                cboResUF.SelectedIndex = 26
                cboComUF.SelectedIndex = 26

            End If
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

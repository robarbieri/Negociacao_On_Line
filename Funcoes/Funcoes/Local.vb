Option Explicit On
Option Strict Off

Imports ConnectTo
Imports System.Math
Imports System.Data.SqlClient

Public Class Local

    Public Function SoNumeros(ByVal strDado As String) As String

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

    Public Function UpperTrim(ByVal strText As String) As String

        Try

            UpperTrim = Trim(UCase(strText))

        Catch ex As Exception

            UpperTrim = ""

        End Try

    End Function

    'Public Function AtualizaValor(ByVal ID_Carteira As Integer, ByVal VencimentoDebito As Date, ByVal AtualizaAte As Date, ByVal Valor As Double, ByVal Sinal As String, ByVal PermiteAtualizacao As Boolean, ByVal QtdParc As Integer, ByVal DataRecebimento As Date, ByVal TDOC_ID As Integer, ByVal DataAtualizacao As Date, ByVal DataVencDebito As Date)
    Public Function AtualizaValor(ByVal ID_Carteira As Integer, ByVal VencimentoDebito As String, ByVal AtualizaAte As String, ByVal Valor As Double, ByVal Sinal As String, ByVal PermiteAtualizacao As Boolean, ByVal QtdParc As Integer, ByVal DataRecebimento As String, ByVal TDOC_ID As Integer, ByVal DataAtualizacao As String, ByVal DataVencDebito As String)
        Dim Conn As New Comando
        Dim DRIGPMHoje As SqlDataReader = Nothing
        Dim DRIGPMVencimento As SqlDataReader = Nothing
        Dim DRPolitica As SqlDataReader = Nothing
        Dim DRTaxaJurosFaixa As SqlDataReader = Nothing
        Dim DRCont As SqlDataReader = Nothing
        Dim vPercentMulta, vMultaPrincipal, vMultaPrincipalCorrigido, vTaxaJuros, vPercentJuros As Double
        Dim vNumDias, vValorCorrigido, vValorMulta, vValorHonorarios, vPercentHonorarios As Double
        Dim vValorJuros As Double
        Dim vTaxaAtualizacao, vNumDias2, vNumDias3 As Double
        Dim vTaxaJurosComposta As Double
        Dim vValorMultaIOF, vValorMultaObrig, vValorIOFObrig As Double
        Dim vCONT_ID, vQtdMeses, vAtraso As Integer
        Dim vNumeroCarteira As String
        Dim sai, vHonorariosPrincipal, vHonorariosPrincipalCorrigido, vHonorarios, vAtualizaValor, vJurosSimples, vJurosComposto, vJurosFaixa, vRecebimentoContrato, vJuros, vVencimentoTitulo, vAtualiza, vAtualizaDebito, vAtualizaCredito, vMulta As Boolean
        Dim vDataNova As String

        vQtdMeses = 0
        vValorJuros = 0
        vValorMulta = 0
        vValorCorrigido = Valor
        sai = False

        'VencimentoDebito = Format(VencimentoDebito, "yyy-MM-dd")
        'AtualizaAte = Format(AtualizaAte, "yyy-MM-dd")
        'DataRecebimento = Format(DataRecebimento, "yyy-MM-dd")
        'DataAtualizacao = Format(DataAtualizacao, "yyy-MM-dd")
        'DataVencDebito = Format(DataVencDebito, "yyy-MM-dd")

        Conn.Banco = "NEOWEB"
        'Conn.Banco = "NEOWEB_REAL"

        DRPolitica = Conn.ExecuteQuery("SELECT * FROM Politica_de_Negociacao p WITH (NOLOCK) JOIN REL_PolNeg_Cart r ON p.PNEG_ID = r.PNEG_ID WHERE CART_ID = " & ID_Carteira)
        DRPolitica.Read()
        DRCont = Conn.ExecuteQuery("SELECT CONT_ID, CART_Numero FROM Carteiras WITH (NOLOCK) WHERE CART_ID = " & ID_Carteira)
        DRCont.Read()
        vCONT_ID = DRCont("CONT_ID")
        vNumeroCarteira = DRCont("CART_Numero")

        If DRPolitica.HasRows Then
            If Trim(DRPolitica("PNEG_Honorarios").ToString) <> "" Then vHonorarios = DRPolitica("PNEG_Honorarios")
            If Trim(DRPolitica("PNEG_PercentHonorarios").ToString) <> "" Then vPercentHonorarios = DRPolitica("PNEG_PercentHonorarios")
            If Trim(DRPolitica("PNEG_HonorariosPrincipal").ToString) <> "" Then vHonorariosPrincipal = DRPolitica("PNEG_HonorariosPrincipal")
            If Trim(DRPolitica("PNEG_HonorariosPrincipalCorrigido").ToString) <> "" Then vHonorariosPrincipalCorrigido = DRPolitica("PNEG_HonorariosPrincipalCorrigido")

            If vCONT_ID = 101 Then 'ID_Carteira = 18 or ID_Carteira = 202 then
                vAtraso = DateDiff("d", DataVencDebito, AtualizaAte)
                If vAtraso <= 150 Then
                    vPercentHonorarios = 10
                ElseIf vAtraso >= 151 And vAtraso <= 360 And QtdParc = 1 Then
                    vPercentHonorarios = 15
                ElseIf vAtraso >= 151 And vAtraso <= 360 And QtdParc > 1 Then
                    vPercentHonorarios = 10
                End If
                'elseif vCONT_ID = 34 or vCONT_ID = 100 then
                '	vAtraso = DateDiff("d", DataVencDebito, AtualizaAte)
                '	if vAtraso <= 60 then
                '		vPercentHonorarios = 0
                '	end if
            End If
        End If

        If PermiteAtualizacao And VencimentoDebito < AtualizaAte Then

            DRPolitica = Conn.ExecuteQuery("SELECT * FROM Politica_de_Negociacao p WITH (NOLOCK) JOIN REL_PolNeg_Cart r ON p.PNEG_ID = r.PNEG_ID WHERE CART_ID = " & ID_Carteira)
            DRPolitica.Read()
            DRCont = Conn.ExecuteQuery("SELECT CONT_ID FROM Carteiras WITH (NOLOCK) WHERE CART_ID = " & ID_Carteira)
            DRCont.Read()

            If DRPolitica.HasRows Then
                If DRPolitica("PNEG_Atualiza").ToString <> "" Then vAtualiza = DRPolitica("PNEG_Atualiza")
                If DRPolitica("PNEG_AtualizaDebito").ToString <> "" Then vAtualizaDebito = DRPolitica("PNEG_AtualizaDebito")
                If DRPolitica("PNEG_AtualizaCredito").ToString <> "" Then vAtualizaCredito = DRPolitica("PNEG_AtualizaCredito")
                If DRPolitica("PNEG_Multa").ToString <> "" Then vMulta = DRPolitica("PNEG_Multa")
                If DRPolitica("PNEG_PercentMulta").ToString <> "" Then vPercentMulta = DRPolitica("PNEG_PercentMulta")
                If DRPolitica("PNEG_MultaPrincipal").ToString <> "" Then vMultaPrincipal = DRPolitica("PNEG_MultaPrincipal")
                If DRPolitica("PNEG_MultaPrincipalCorrigido").ToString <> "" Then vMultaPrincipalCorrigido = DRPolitica("PNEG_MultaPrincipalCorrigido")
                If DRPolitica("PNEG_AtualizaValor").ToString <> "" Then vAtualizaValor = DRPolitica("PNEG_AtualizaValor")
                If DRPolitica("TAXA_ID_Atualizacao").ToString <> "" Then vTaxaAtualizacao = DRPolitica("TAXA_ID_Atualizacao")
                If DRPolitica("PNEG_Juros").ToString <> "" Then vJuros = DRPolitica("PNEG_Juros")
                If DRPolitica("PNEG_JurosPorFaixa").ToString <> "" Then vJurosFaixa = DRPolitica("PNEG_JurosPorFaixa")
                If DRPolitica("PNEG_PercentualJuros").ToString <> "" Then vPercentJuros = DRPolitica("PNEG_PercentualJuros")
                If DRPolitica("PNEG_JurosSimples").ToString <> "" Then vJurosSimples = DRPolitica("PNEG_JurosSimples")
                If DRPolitica("PNEG_JurosComposto").ToString <> "" Then vJurosComposto = DRPolitica("PNEG_JurosComposto")
                If DRPolitica("PNEG_Honorarios").ToString <> "" Then vHonorarios = DRPolitica("PNEG_Honorarios")
                If DRPolitica("PNEG_PercentHonorarios").ToString <> "" Then vPercentHonorarios = DRPolitica("PNEG_PercentHonorarios")
                If DRPolitica("PNEG_HonorariosPrincipal").ToString <> "" Then vHonorariosPrincipal = DRPolitica("PNEG_HonorariosPrincipal")
                If DRPolitica("PNEG_HonorariosPrincipalCorrigido").ToString <> "" Then vHonorariosPrincipalCorrigido = DRPolitica("PNEG_HonorariosPrincipalCorrigido")
                If DRPolitica("PNEG_AtualizaDoVencimento").ToString <> "" Then vVencimentoTitulo = DRPolitica("PNEG_AtualizaDoVencimento")
                If DRPolitica("PNEG_AtualizaDoRecebimento").ToString <> "" Then vRecebimentoContrato = DRPolitica("PNEG_AtualizaDoRecebimento")

                If vCONT_ID = 101 Then 'ID_Carteira = 18 or ID_Carteira = 202 then
                    vAtraso = DateDiff("d", DataVencDebito, AtualizaAte)
                    If vAtraso <= 150 Then
                        vPercentHonorarios = 10
                    ElseIf vAtraso >= 151 And vAtraso <= 360 And QtdParc = 1 Then
                        vPercentHonorarios = 15
                    ElseIf vAtraso >= 151 And vAtraso <= 360 And QtdParc > 1 Then
                        vPercentHonorarios = 10
                    End If
                    'elseif vCONT_ID = 34 or vCONT_ID = 100 then
                    '	vAtraso = DateDiff("d", DataVencDebito, AtualizaAte)
                    '	if vAtraso <= 60 then
                    '		vPercentHonorarios = 0
                    '	end if
                End If

                If (ID_Carteira = 41 Or ID_Carteira = 42 Or ID_Carteira = 43 Or ID_Carteira = 44) And AtualizaAte > CDate("17/10/2005") Then
                    AtualizaAte = CDate("17/10/2005")
                End If
                If vCONT_ID = 43 And AtualizaAte > CDate("10/04/2006") Then
                    AtualizaAte = CDate("10/04/2006")
                End If

                If vRecebimentoContrato Then
                    vNumDias = DateDiff("d", DataRecebimento, AtualizaAte)
                Else
                    vNumDias = DateDiff("d", VencimentoDebito, AtualizaAte)
                End If

                vNumDias2 = DateDiff("d", DataVencDebito, AtualizaAte)

                Dim DRTaxaHoje, DRTaxaVenc As SqlDataReader
                Dim vNumMeses, vDiasRestantes, w As Integer
                Dim vDescTipoDoc As String
                Dim DRDescTipoDoc As SqlDataReader = Nothing

                vDescTipoDoc = ""
                If DRCont("CONT_ID") = 10 Or DRCont("CONT_ID") = 31 Then
                    DRDescTipoDoc = Conn.ExecuteQuery("SELECT TDOC_Descricao FROM Tipos_de_Documento WITH (NOLOCK) WHERE TDOC_ID = " & TDOC_ID)
                    DRDescTipoDoc.Read()
                    If DRDescTipoDoc.HasRows Then
                        vDescTipoDoc = DRDescTipoDoc("TDOC_Descricao")
                    End If
                    DRDescTipoDoc.Close()
                    DRDescTipoDoc = Nothing
                End If

                If DRCont("CONT_ID") = 83 Then
                    vNumDias3 = vNumDias
                    If vNumDias3 > 114 Then
                        vNumDias3 = 114
                    End If
                    vValorJuros = (vValorCorrigido * (1 + (12.9 / 100)) ^ (vNumDias3 / 30)) - vValorCorrigido
                    vValorCorrigido = vValorCorrigido + vValorJuros
                ElseIf DRCont("CONT_ID") = 10 Then
                    If vDescTipoDoc = "ADIANTAMENTO DEPOSITANTE" Or vDescTipoDoc = "CHEQUE ESPECIAL" Or vDescTipoDoc = "CHEQUE UNIVERSITARIO MB" Or vDescTipoDoc = "CHEQUE EMPRESA MB" Then
                        If vNumDias < 180 Then
                            vPercentJuros = 7
                        Else
                            vPercentJuros = 8
                        End If
                    Else
                        If vNumDias < 180 Then
                            vPercentJuros = 6
                        Else
                            vPercentJuros = 7
                        End If
                    End If
                    vValorJuros = (vValorCorrigido * (1 + (vPercentJuros / 100)) ^ (vNumDias / 30)) - vValorCorrigido
                    vValorCorrigido = vValorCorrigido + vValorJuros
                ElseIf DRCont("CONT_ID") = 99 Then
                    vValorCorrigido = Round(Valor * (1.0002125 ^ vNumDias), 2)
                    vValorMulta = Round(vValorCorrigido * (vPercentMulta / 100), 2)
                    vValorJuros = Round((Valor + vValorMulta) * ((vNumDias * (vPercentJuros / 30)) / 100), 2)
                    vValorCorrigido = vValorCorrigido + vValorJuros + vValorMulta
                ElseIf DRCont("CONT_ID") = 50 Then
                    If vNumDias < 180 Then
                        vValorMulta = vValorCorrigido * 0.02
                        vNumMeses = vNumDias \ 30
                        vDiasRestantes = vNumDias Mod 30
                        If vNumMeses > 0 Then
                            For w = 1 To vNumMeses
                                vValorCorrigido = vValorCorrigido * 1.099 ' 9,9% ao mês
                            Next
                        End If
                        If vDiasRestantes > 0 Then
                            vValorCorrigido = vValorCorrigido * (1 + (9.9 / 3000 * vDiasRestantes))
                        End If
                        vValorCorrigido = vValorCorrigido + vValorMulta
                    Else
                        DRTaxaVenc = Conn.ExecuteQuery("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & "01/" & Month(VencimentoDebito) & "/" & Year(VencimentoDebito) & "') AND TAXA_ID = 10")
                        DRTaxaVenc.Read()

                        If Not DRTaxaVenc.HasRows Then
                            DRTaxaVenc = Conn.ExecuteQuery("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & "01/" & Month(VencimentoDebito) & "/" & Year(VencimentoDebito) & "') AND TAXA_ID = 10 ORDER BY COTA_Data DESC")
                            DRTaxaVenc.Read()
                        Else
                            vValorCorrigido = vValorCorrigido * (1 + DRTaxaVenc("COTA_Indice"))
                        End If

                        DRTaxaVenc.Close()
                    End If
                ElseIf TDOC_ID = 67 And DateDiff("d", VencimentoDebito, AtualizaAte) < 61 Then
                    'Mora
                    DRTaxaHoje = Conn.ExecuteQuery("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 4")
                    DRTaxaVenc = Conn.ExecuteQuery("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 4")
                    DRTaxaHoje.Read()
                    DRTaxaVenc.Read()

                    If Not DRTaxaHoje.HasRows Then
                        DRTaxaHoje = Conn.ExecuteQuery("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 4 ORDER BY COTA_Data DESC")
                        DRTaxaHoje.Read()
                    End If

                    If Not DRTaxaVenc.HasRows Then
                        DRTaxaVenc = Conn.ExecuteQuery("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 4 ORDER BY COTA_Data DESC")
                        DRTaxaVenc.Read()
                    End If

                    If DRTaxaHoje.HasRows And DRTaxaVenc.HasRows Then
                        vValorCorrigido = (DRTaxaHoje("COTA_Indice") / DRTaxaVenc("COTA_Indice")) * vValorCorrigido
                    End If

                    DRTaxaVenc.Close()
                    DRTaxaHoje.Close()
                ElseIf TDOC_ID = 67 And DateDiff("d", VencimentoDebito, AtualizaAte) > 60 And DateDiff("d", VencimentoDebito, DataAtualizacao) < 61 Then
                    'Mora
                    DRTaxaHoje = Conn.ExecuteQuery("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & DateAdd("d", 60, VencimentoDebito) & "') AND TAXA_ID = 4")
                    DRTaxaVenc = Conn.ExecuteQuery("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 4")
                    DRTaxaHoje.Read()
                    DRTaxaVenc.Read()

                    If Not DRTaxaHoje.HasRows Then
                        DRTaxaHoje = Conn.ExecuteQuery("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & DateAdd("d", 60, VencimentoDebito) & "') AND TAXA_ID = 4 ORDER BY COTA_Data DESC")
                        DRTaxaHoje.Read()
                    End If

                    If Not DRTaxaVenc.HasRows Then
                        DRTaxaVenc = Conn.ExecuteQuery("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 4 ORDER BY COTA_Data DESC")
                        DRTaxaVenc.Read()
                    End If

                    If DRTaxaHoje.HasRows And DRTaxaVenc.HasRows Then
                        vValorCorrigido = (DRTaxaHoje("COTA_Indice") / DRTaxaVenc("COTA_Indice")) * vValorCorrigido
                    End If

                    DRTaxaVenc.Close()
                    DRTaxaHoje.Close()

                    'CL/LP
                    DRTaxaHoje = Conn.ExecuteQuery("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 5")
                    DRTaxaVenc = Conn.ExecuteQuery("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & DateAdd("d", 60, VencimentoDebito) & "') AND TAXA_ID = 5")
                    DRTaxaHoje.Read()
                    DRTaxaVenc.Read()

                    If Not DRTaxaHoje.HasRows Then
                        DRTaxaHoje = Conn.ExecuteQuery("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 5 ORDER BY COTA_Data DESC")
                        DRTaxaHoje.Read()
                    End If

                    If Not DRTaxaVenc.HasRows Then
                        DRTaxaVenc = Conn.ExecuteQuery("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & DateAdd("d", 60, VencimentoDebito) & "') AND TAXA_ID = 5 ORDER BY COTA_Data DESC")
                        DRTaxaVenc.Read()
                    End If

                    If DRTaxaHoje.HasRows And DRTaxaVenc.HasRows Then
                        vValorCorrigido = (DRTaxaHoje("COTA_Indice") / DRTaxaVenc("COTA_Indice")) * vValorCorrigido
                    End If

                    DRTaxaVenc.Close()
                    DRTaxaHoje.Close()
                    DRTaxaVenc = Nothing
                    DRTaxaHoje = Nothing
                    DRTaxaVenc = Nothing
                    DRTaxaHoje = Nothing
                ElseIf (TDOC_ID = 67 And DateDiff("d", VencimentoDebito, AtualizaAte) > 60 And DateDiff("d", VencimentoDebito, DataAtualizacao) > 60) Or TDOC_ID = 68 Or TDOC_ID = 69 Then
                    'CL/LP
                    DRTaxaHoje = Conn.ExecuteQuery("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 5")
                    DRTaxaVenc = Conn.ExecuteQuery("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 5")
                    DRTaxaHoje.Read()
                    DRTaxaVenc.Read()

                    If Not DRTaxaHoje.HasRows Then
                        DRTaxaHoje = Conn.ExecuteQuery("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 5 ORDER BY COTA_Data DESC")
                        DRTaxaHoje.Read()
                    End If

                    If Not DRTaxaVenc.HasRows Then
                        DRTaxaVenc = Conn.ExecuteQuery("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & DataAtualizacao & "') AND TAXA_ID = 5 ORDER BY COTA_Data DESC")
                        DRTaxaVenc.Read()
                    End If

                    If DRTaxaHoje.HasRows And DRTaxaVenc.HasRows Then
                        vValorCorrigido = (DRTaxaHoje("COTA_Indice") / DRTaxaVenc("COTA_Indice")) * vValorCorrigido
                    End If

                    DRTaxaVenc.Close()
                    DRTaxaHoje.Close()
                    DRTaxaVenc = Nothing
                    DRTaxaHoje = Nothing
                Else
                    If vAtualiza Then
                        If (vAtualizaDebito And Sinal = "+") Or (vAtualizaCredito And Sinal = "-") Then
                            If vJuros Then
                                vValorJuros = 0
                                'Cobrar juros
                                If vJurosFaixa Then
                                    'Juros por faixa de atraso
                                    'Set DRTaxaJurosFaixa = Conn.Executequery("SELECT FNEG_PercentJuros, FNEG_JurosSimples, FNEG_JurosComposto FROM Faixas_de_Negociacao f WITH (NOLOCK) JOIN Tabelas_de_Negociacao t WITH (NOLOCK) ON f.TNEG_ID = t.TNEG_ID JOIN REL_TabNeg_Cart r WITH (NOLOCK) ON t.TNEG_ID = r.TNEG_ID WHERE " & vNumDias2 & " BETWEEN TNEG_InicioFaixaAtraso AND TNEG_FinalFaixaAtraso AND CART_ID = " & ID_Carteira & " AND FNEG_QtdParcelas = " & QtdParc & " AND FNEG_PercentJuros IS NOT NULL AND FNEG_Habilitado = 1")
                                    DRTaxaJurosFaixa = Conn.ExecuteQuery("SELECT FNEG_PercentJuros, FNEG_JurosSimples, FNEG_JurosComposto FROM Faixas_de_Negociacao f WITH (NOLOCK) JOIN Tabelas_de_Negociacao t WITH (NOLOCK) ON f.TNEG_ID = t.TNEG_ID JOIN REL_TabNeg_Cart r WITH (NOLOCK) ON t.TNEG_ID = r.TNEG_ID WHERE ((" & vNumDias2 & " BETWEEN TNEG_InicioFaixaAtraso AND TNEG_FinalFaixaAtraso) OR (" & Year(DataVencDebito) & " BETWEEN TNEG_InicioAnoAtraso AND TNEG_FinalAnoAtraso)) AND CART_ID = " & ID_Carteira & " AND FNEG_QtdParcelas = " & QtdParc & " AND FNEG_PercentJuros IS NOT NULL AND FNEG_Habilitado = 1 AND TNEG_Habilitada = 1 AND TNEG_VigenciaDe <= '" & AtualizaAte & "' AND TNEG_VigenciaAte >= '" & AtualizaAte & "'")
                                    If DRTaxaJurosFaixa.Read Then
                                        vJurosSimples = DRTaxaJurosFaixa("FNEG_JurosSimples")
                                        vJurosComposto = DRTaxaJurosFaixa("FNEG_JurosComposto")
                                        vPercentJuros = DRTaxaJurosFaixa("FNEG_PercentJuros")

                                        If vJurosSimples Then
                                            If DRCont("CONT_ID") = 31 Then
                                                vValorJuros = (((vNumDias * Round(vPercentJuros / 30, 3)) / 100) * vValorCorrigido)
                                            Else
                                                vValorJuros = (((vNumDias * (vPercentJuros / 30)) / 100) * vValorCorrigido)
                                            End If
                                        ElseIf vJurosComposto Then
                                            vValorJuros = (vValorCorrigido * (1 + (vPercentJuros / 100)) ^ (vNumDias / 30)) - vValorCorrigido
                                        End If
                                    End If
                                    DRTaxaJurosFaixa.Close()
                                    DRTaxaJurosFaixa = Nothing
                                ElseIf vPercentJuros <> 0 Then
                                    If vJurosSimples Then
                                        If DRCont("CONT_ID") = 31 Then
                                            vValorJuros = (((vNumDias * Round(vPercentJuros / 30, 3)) / 100) * vValorCorrigido)
                                        Else
                                            vValorJuros = (((vNumDias * (vPercentJuros / 30)) / 100) * vValorCorrigido)
                                        End If
                                    ElseIf vJurosComposto Then
                                        vValorJuros = (vValorCorrigido * (1 + (vPercentJuros / 100)) ^ (vNumDias / 30)) - vValorCorrigido
                                    End If
                                End If
                            End If
                            If vAtualizaValor And UCase(vDescTipoDoc) <> "CREDUCSAL" Then
                                If vTaxaAtualizacao = 3 Then
                                    'IGPM
                                    DRIGPMHoje = Conn.ExecuteQuery("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 3")
                                    DRIGPMVencimento = Conn.ExecuteQuery("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & VencimentoDebito & "') + 2 AND TAXA_ID = 3")
                                    DRIGPMHoje.Read()
                                    DRIGPMVencimento.Read()

                                    If Not DRIGPMHoje.HasRows Then
                                        DRIGPMHoje = Conn.ExecuteQuery("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = 3 ORDER BY COTA_Data DESC")
                                        DRIGPMHoje.Read()
                                    End If

                                    If Not DRIGPMVencimento.HasRows Then
                                        DRIGPMVencimento = Conn.ExecuteQuery("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & VencimentoDebito & "') AND TAXA_ID = 3 ORDER BY COTA_Data DESC")
                                        DRIGPMVencimento.Read()
                                    End If

                                    If Not DRIGPMHoje.HasRows Then
                                        DRIGPMHoje = Conn.ExecuteQuery("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE TAXA_ID = 3 ORDER BY COTA_Data DESC")
                                        DRIGPMHoje.Read()
                                    End If

                                    If Not DRIGPMVencimento.HasRows Then
                                        DRIGPMVencimento = Conn.ExecuteQuery("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE TAXA_ID = 3 ORDER BY COTA_Data DESC")
                                        DRIGPMVencimento.Read()
                                    End If

                                    'If Trim(DRIGPMHoje("COTA_Indice").ToString) <> "" And Trim(DRIGPMVencimento("COTA_Indice")).ToString <> "" Then
                                    vValorCorrigido = (DRIGPMHoje("COTA_Indice") / DRIGPMVencimento("COTA_Indice")) * vValorCorrigido
                                    'End If

                                    DRIGPMVencimento.Close()
                                    DRIGPMHoje.Close()
                                    DRIGPMVencimento = Nothing
                                    DRIGPMHoje = Nothing
                                ElseIf vTaxaAtualizacao = 10 Then
                                    'Taxa Renner
                                    DRIGPMHoje = Conn.ExecuteQuery("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = " & vTaxaAtualizacao)
                                    DRIGPMVencimento = Conn.ExecuteQuery("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & VencimentoDebito & "') AND TAXA_ID = " & vTaxaAtualizacao)
                                    DRIGPMHoje.Read()
                                    DRIGPMVencimento.Read()

                                    If Not DRIGPMHoje.HasRows Then
                                        DRIGPMHoje = Conn.ExecuteQuery("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
                                        DRIGPMHoje.Read()
                                    End If

                                    If Not DRIGPMVencimento.HasRows Then
                                        DRIGPMVencimento = Conn.ExecuteQuery("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & VencimentoDebito & "') AND TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
                                        DRIGPMVencimento.Read()
                                        If Not DRIGPMVencimento.HasRows Then
                                            DRIGPMVencimento = Conn.ExecuteQuery("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data")
                                            DRIGPMVencimento.Read()
                                        End If
                                    End If

                                    'if DRIGPMHoje.EOF or DRIGPMVencimento.EOF then
                                    '	vValorCorrigido = Valor
                                    'else
                                    vValorCorrigido = (DRIGPMVencimento("COTA_Indice") / DRIGPMHoje("COTA_Indice")) * vValorCorrigido
                                    'end if

                                    DRIGPMVencimento.Close()
                                    DRIGPMHoje.Close()
                                    DRIGPMVencimento = Nothing
                                    DRIGPMHoje = Nothing
                                ElseIf vTaxaAtualizacao = 15 Then
                                    'IGPM Pró-rata
                                    If CDate(AtualizaAte) > CDate(VencimentoDebito) Then
                                        If Month(VencimentoDebito) = Month(AtualizaAte) And Year(VencimentoDebito) = Year(AtualizaAte) Then
                                            DRIGPMVencimento = Conn.ExecuteQuery("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & "01/" & Month(VencimentoDebito) & "/" & Year(VencimentoDebito) & "') AND TAXA_ID = " & vTaxaAtualizacao)
                                            DRIGPMVencimento.Read()
                                            If DRIGPMVencimento.HasRows Then
                                                vValorCorrigido = DateDiff("d", VencimentoDebito, AtualizaAte) * DRIGPMVencimento("COTA_Indice") / 100 / 30 * vValorCorrigido
                                            End If
                                        Else
                                            sai = False
                                            vTaxaJurosComposta = 0
                                            DRIGPMVencimento = Conn.ExecuteQuery("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & "01/" & Month(VencimentoDebito) & "/" & Year(VencimentoDebito) & "') AND TAXA_ID = " & vTaxaAtualizacao)
                                            DRIGPMVencimento.Read()
                                            If DRIGPMVencimento.HasRows Then
                                                vTaxaJurosComposta = 1 + Round((DateDiff("d", VencimentoDebito, DateAdd("m", 1, ("01/" & Month(VencimentoDebito) & "/" & Year(VencimentoDebito)))) * DRIGPMVencimento("COTA_Indice") / 100 / 30), 6)
                                            End If
                                            Dim y As Date
                                            y = VencimentoDebito
                                            Do While Not sai
                                                y = DateAdd("m", 1, y)
                                                If DateDiff("m", w, AtualizaAte) = 0 Then
                                                    sai = True
                                                    DRIGPMVencimento = Conn.ExecuteQuery("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & "01/" & Month(AtualizaAte) & "/" & Year(AtualizaAte) & "') AND TAXA_ID = " & vTaxaAtualizacao)
                                                    DRIGPMVencimento.Read()
                                                    If DRIGPMVencimento.HasRows Then
                                                        vTaxaJurosComposta = Round(vTaxaJurosComposta * Round((1 + (DateDiff("d", "01/" & Month(AtualizaAte) & "/" & Year(AtualizaAte), AtualizaAte) * DRIGPMVencimento("COTA_Indice") / 100 / 30)), 6), 6)
                                                    End If
                                                Else
                                                    DRIGPMVencimento = Conn.ExecuteQuery("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & "01/" & Month(y) & "/" & Year(y) & "') AND TAXA_ID = " & vTaxaAtualizacao)
                                                    DRIGPMVencimento.Read()
                                                    If DRIGPMVencimento.HasRows Then
                                                        vTaxaJurosComposta = Round(vTaxaJurosComposta * Round((1 + DRIGPMVencimento("COTA_Indice") / 100), 6), 6)
                                                    End If
                                                End If
                                            Loop
                                            vValorCorrigido = vValorCorrigido + Round((vTaxaJurosComposta - 1) * vValorCorrigido, 2)
                                        End If
                                    End If
                                Else
                                    'Qualquer outro índice
                                    DRIGPMHoje = Conn.ExecuteQuery("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = " & vTaxaAtualizacao)
                                    DRIGPMVencimento = Conn.ExecuteQuery("SELECT * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data = convert(smalldatetime,'" & VencimentoDebito & "') AND TAXA_ID = " & vTaxaAtualizacao)
                                    DRIGPMHoje.Read()
                                    DRIGPMVencimento.Read()

                                    If Not DRIGPMHoje.HasRows Then
                                        DRIGPMHoje = Conn.ExecuteQuery("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & AtualizaAte & "') AND TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
                                        DRIGPMHoje.Read()
                                    End If

                                    If Not DRIGPMVencimento.HasRows Then
                                        DRIGPMVencimento = Conn.ExecuteQuery("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE COTA_Data < convert(smalldatetime,'" & VencimentoDebito & "') AND TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data DESC")
                                        DRIGPMVencimento.Read()
                                        If Not DRIGPMVencimento.HasRows Then
                                            DRIGPMVencimento = Conn.ExecuteQuery("SELECT TOP 1 * FROM Cotacoes WITH (NOLOCK) WHERE TAXA_ID = " & vTaxaAtualizacao & " ORDER BY COTA_Data")
                                            DRIGPMVencimento.Read()
                                        End If
                                    End If

                                    vValorCorrigido = (DRIGPMHoje("COTA_Indice") / DRIGPMVencimento("COTA_Indice")) * vValorCorrigido

                                    DRIGPMVencimento.Close()
                                    DRIGPMHoje.Close()
                                    DRIGPMVencimento = Nothing
                                    DRIGPMHoje = Nothing
                                End If
                                If DRCont("CONT_ID") = 31 Then
                                    vValorJuros = (((vNumDias * Round(vPercentJuros / 30, 3)) / 100) * vValorCorrigido)
                                End If
                            End If
                            vValorCorrigido = vValorCorrigido + vValorJuros
                            If vMulta Then
                                'Cobrar multa
                                vValorMulta = 0
                                If vCONT_ID = 13 And vNumDias > 120 Then
                                    vPercentMulta = 0
                                End If
                                If vMultaPrincipal Then
                                    'Cobrar multa sobre o principal
                                    'vValorMulta = ((vNumDias * (vPercentMulta / 30)) / 100) * Valor
                                    vValorMulta = (vPercentMulta / 100) * Valor
                                ElseIf vMultaPrincipalCorrigido Then
                                    'Cobrar multa sobre o proncipal corrigido (com Juros)
                                    'vValorMulta = ((vNumDias * (vPercentMulta / 30)) / 100) * vValorCorrigido
                                    vValorMulta = (vPercentMulta / 100) * vValorCorrigido
                                End If
                                If DRCont("CONT_ID") = 31 Then
                                    If UCase(vDescTipoDoc) = "CREDUCSAL" Then
                                        vValorMulta = (vPercentMulta / 100) * Valor
                                    Else
                                        vValorCorrigido = vValorCorrigido - vValorJuros
                                        vValorMulta = (vPercentMulta / 100) * vValorCorrigido
                                        vValorCorrigido = vValorCorrigido + vValorJuros
                                    End If
                                End If
                                vValorCorrigido = vValorCorrigido + vValorMulta
                            End If
                        End If
                    End If
                End If
            End If
        End If

        If vCONT_ID = 1 Then
            vValorMultaObrig = Valor * 0.02
            vValorIOFObrig = (Valor + vValorMultaObrig) * 0.01
            vValorMultaIOF = vValorMultaObrig + vValorIOFObrig
            If vValorMultaIOF > vValorCorrigido - Valor Then
                vValorCorrigido = Valor + vValorMultaIOF
            End If
        End If

        If vHonorarios Then
            'Cobrar honorarios
            vValorHonorarios = 0
            If vHonorariosPrincipal Then
                'Cobrar honorarios sobre o principal
                vValorHonorarios = (vPercentHonorarios / 100) * Valor
            ElseIf vHonorariosPrincipalCorrigido Then
                'Cobrar honorarios sobre o proncipal corrigido (com Juros e Multa)
                vValorHonorarios = (vPercentHonorarios / 100) * vValorCorrigido
            End If
            vValorCorrigido = vValorCorrigido + vValorHonorarios
        End If

        AtualizaValor = vValorCorrigido

    End Function

End Class

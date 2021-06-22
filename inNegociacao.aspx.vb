Option Explicit On
Option Strict Off

Imports System
Imports System.Text
Imports System.Math
Imports System.Data.SqlClient
Imports ConnectTo
Imports XMail

Partial Class inNegociacao
    Inherits System.Web.UI.Page

    Protected dblSaldo As Double = 0
    Protected dblPrincipal As Double = 0
    Protected dblEntradaMin As Double = 0
    Protected eMail As New XMail.SendMail
    Protected strContrato As String = ""
    Protected idCliente As String = ""
    Protected idLogin As Long = 0
    Protected Shared m_executingPages As Hashtable = New Hashtable

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Try
            'Ajax.Utility.RegisterTypeForAjax(GetType(inNegociacao))
            Page.ClientScript.RegisterClientScriptBlock(Me.GetType, "clientScript", "<script language=javascript src=Funcoes/funcoes.js></script>")
            strContrato = Trim(Request.QueryString("ctra"))
            idCliente = Trim(Request.QueryString("idcliente"))
            idLogin = Trim(Request.QueryString("id"))
            'txtidLogin.Text = idLogin
            'txtContratoOS.Text = strContrato
            btnCalcularServer.Attributes.Add("onclick", " this.disabled = true; " + " btnNegociacao.disabled = true; " + _
            ClientScript.GetPostBackEventReference(btnCalcularServer, "") + ";")
            btnNegociacao.Attributes.Add("onclick", " this.disabled = true; " + " btnCalcularServer.disabled = true; " + _
            ClientScript.GetPostBackEventReference(btnNegociacao, "") + ";")
            Try
                If CInt(Trim(Session("Parc"))) <> CInt(Trim(Left(cboCondicao.SelectedItem.Text, 2))) Then
                    Session("Parc") = ""
                    Session("EntradaMin") = ""
                End If
            Catch
            End Try
            If Trim(Session("EntradaMin")) <> "" Then dblEntradaMin = Session("EntradaMin")
            If Trim(Session("Saldo")) = "" Then Call Carregar()
        Catch ex As Exception
            SendMailErro("Erro na Negociação On-Line." & Chr(13) & Chr(13) & "Número: " & Err.Number & Chr(13) & "Source: " & Err.Source & Chr(13) & "Descrição: " & Err.Description & Chr(13) & "Help: " & Chr(13) & Err.HelpFile & Chr(13) & Err.HelpContext & Chr(13) & Err.LastDllError)
            ShowMessage("Sistema Inativo. Por favor tente novamente em alguns minutos ou entre em contato com nosso atendimento no telefone 11-2171-0301(SP) ou 0800-6000-500(Demais Localidades).")
            'Response.Redirect("inCPF.aspx")
        End Try

    End Sub

    Protected Sub Carregar()

        Dim intAtraso As Integer = 0
        Dim idCTRA As Long = 0
        Dim Conn As New Comando
        Dim objDR As SqlDataReader = Nothing
        Dim objDRMirror As SqlDataReader = Nothing
        Dim objDRDOC As SqlDataReader = Nothing
        Dim strSql As String = ""
        Dim x As Integer = 0
        Dim strTabelas As String = ""
        'Dim arrFaixas(100) As Integer
        Dim datVenctoContrato As Date
        Dim datRecebimento As Date
        'Dim Funcoes As New Funcoes.Local

        Try

            Conn.Banco = "NEOWEB"
            'Conn.Banco = "NEOWEB_REAL"

            dblEntradaMin = 0

            strSql = "Select CTRA_Id,DateDiff(Day,CTRA_VencDebito,GetDate()) as Atraso, CTRA_VencDebito as Vencto, CTRA_DataRecebimentoContrato as Recebimento " & _
                     "From Contratos (NOLOCK) Where CTRA_Numero = '" & strContrato & "'"
            objDR = Conn.ExecuteQuery(strSql)

            objDR.Read()
            idCTRA = CLng(objDR("CTRA_Id"))
            intAtraso = CInt(objDR("Atraso"))
            datVenctoContrato = CDate(objDR("Vencto"))
            datRecebimento = CDate(objDR("Recebimento"))

            Conn.Banco = "MIRRORWEB"

            strSql = "Select TOP 1 FNEG_Honorarios from tb_Faixas_de_NegociacaoOnLine Where FNEG_Campanha = '" & Trim(Request.QueryString("cod")) & "'"
            objDRMirror = Conn.ExecuteQuery(strSql)

            If objDRMirror.Read Then
                Session("honor") = Trim(objDRMirror("FNEG_Honorarios"))
            Else
                Session("honor") = "20"
                SendMailErro("Percentual de Honorários nãoi localizado para campanha: " & Trim(Request.QueryString("cod")) & Chr(13) & "Query: " & strSql)
            End If

            Conn.Banco = "NEOWEB"
            'Conn.Banco = "NEOWEB_REAL"

            strSql = "SELECT t.TRAN_Valor as Saldo, td.TDOC_Id as Id, t.TRAN_Vencimento as Vencto, t.TRAN_DataRecebimento as Recebimento " & _
                     "FROM Transacoes t WITH (NOLOCK) LEFT JOIN Tipos_de_Documento td WITH (NOLOCK) " & _
                     "ON t.TDOC_ID = td.TDOC_ID " & _
                     "WHERE CTRA_ID = " & idCTRA & _
                     " ORDER BY TRAN_NumTitulo, TRAN_Vencimento"
            '"AND TRAN_Vencimento <= '" & Format(CDate(objDR("Vencto")), "yyy-MM-dd") & "' ORDER BY TRAN_NumTitulo, TRAN_Vencimento"
            objDR = Conn.ExecuteQuery(strSql)

            dblSaldo = 0
            dblPrincipal = 0
            Session("PercentHonor") = ""
            Session("valCorrecao") = ""

            Do While objDR.Read()
                If objDR("Id").ToString = "" Or objDR("Id").ToString = "0" Then
                    'dblSaldo = dblSaldo + CDbl(objDR("Saldo"))
                    dblSaldo = dblSaldo + AtualizaValor(CInt(Session("Carteira")), Format(CDate(Trim(objDR("Vencto"))), "yyy-MM-dd"), Format(CDate(Now), "yyy-MM-dd"), CDbl(objDR("Saldo")), "+", True, 1, Format(CDate(datRecebimento), "yyy-MM-dd"), -1, Format(CDate(Trim(objDR("Recebimento"))), "yyy-MM-dd"), Format(CDate(datVenctoContrato), "yyy-MM-dd"))
                    dblPrincipal = dblPrincipal + CDbl(objDR("Saldo"))
                    'vValorAtualizadoVenc = AtualizaValor(vCART_ID, RSTransacoes("TRAN_Vencimento"), vData, ValorPrincipal(ais("BAND_ID")), "+", 1, 1, vDataRecebimento, -1, RSTransacoes("TRAN_DataRecebimento"), RSDadosContratuais("CTRA_VencDebito"))
                Else
                    strSql = "Select TOP 1 TDOC_Sinal, TDOC_PermiteAtualizacao From Tipos_de_Documento Where TDOC_Id = " & CInt(objDR("Id"))
                    objDRDOC = Conn.ExecuteQuery(strSql)
                    objDRDOC.Read()

                    If Trim(objDRDOC("TDOC_Sinal")) = "+" Then
                        dblSaldo = dblSaldo + AtualizaValor(CInt(Session("Carteira")), Format(CDate(Trim(objDR("Vencto"))), "yyy-MM-dd"), Format(CDate(Now), "yyy-MM-dd"), CDbl(objDR("Saldo")), Trim(objDRDOC("TDOC_Sinal")), objDRDOC("TDOC_PermiteAtualizacao"), 1, Format(CDate(datRecebimento), "yyy-MM-dd"), CInt(objDR("Id")), Format(CDate(Trim(objDR("Recebimento"))), "yyy-MM-dd"), Format(CDate(datVenctoContrato), "yyy-MM-dd"))
                        dblPrincipal = dblPrincipal + CDbl(objDR("Saldo"))
                    Else
                        dblSaldo = dblSaldo - AtualizaValor(CInt(Session("Carteira")), Format(CDate(Trim(objDR("Vencto"))), "yyy-MM-dd"), Format(CDate(Now), "yyy-MM-dd"), CDbl(objDR("Saldo")), Trim(objDRDOC("TDOC_Sinal")), objDRDOC("TDOC_PermiteAtualizacao"), 1, Format(CDate(datRecebimento), "yyy-MM-dd"), CInt(objDR("Id")), Format(CDate(Trim(objDR("Recebimento"))), "yyy-MM-dd"), Format(CDate(datVenctoContrato), "yyy-MM-dd"))
                        dblPrincipal = dblPrincipal - CDbl(objDR("Saldo"))
                    End If
                End If

            Loop

            dblPrincipal = Math.Round(dblPrincipal, 2)
            dblSaldo = Math.Round(dblSaldo, 2)

            strSql = "Select A.TNEG_Id From Tabelas_de_Negociacao A (NOLOCK) Join " & _
                     "REL_TabNeg_Cart B (NOLOCK) On B.TNEG_Id = A.TNEG_Id " & _
                     "Where " & intAtraso & " Between A.TNEG_InicioFaixaAtraso and A.TNEG_FinalFaixaAtraso AND " & _
                     Replace(dblSaldo, ",", ".") & " Between A.TNEG_InicioFaixaSaldo and A.TNEG_FinalFaixaSaldo AND " & _
                     "B.CART_Id = " & Session("Carteira")
            objDR = Conn.ExecuteQuery(strSql)

            Do While objDR.Read

                If x > 0 Then strTabelas = strTabelas & ","
                strTabelas = strTabelas & Trim(objDR("TNEG_Id"))
                x = x + 1

            Loop

            x = 0

            strSql = "Select FNEG_Id, FNEG_QtdParcelas, FNEG_descprincipal, FNEG_desccorrecao From Faixas_de_Negociacao (NOLOCK) Where " & _
                     "TNEG_Id In(" & strTabelas & ") AND FNEG_Habilitado = 1 Order By FNEG_QtdParcelas, FNEG_desccorrecao, FNEG_descprincipal"
            objDR = Conn.ExecuteQuery(strSql)

            Conn.Banco = "MIRRORWEB"

            'strSql = "Select FNEG_QtdParcelas,FNEG_DescPrincipal From tb_Faixas_de_NegociacaoOnLine (NOLOCK) " & _
            '             "Where FNEG_Campanha = '" & Trim(Request.QueryString("cod")) & "' Order By FNEG_QtdParcelas"
            'objDRMirror = Conn.ExecuteQuery(strSql)

            cboCondicao.Items.Clear()
            cboCondicao.Items.Add("Selecione...")

            Do While objDR.Read
                'x = 0
                'Do While objDRMirror.Read

                strSql = "Select TOP 1 FNEG_QtdParcelas,FNEG_DescPrincipal From tb_Faixas_de_NegociacaoOnLine (NOLOCK) " & _
                         "Where FNEG_Id = " & Trim(objDR("FNEG_Id")) & " AND FNEG_Campanha = '" & Trim(Request.QueryString("cod")) & "' "
                If intAtraso < 180 And Session("Carteira") = "87001" Then
                Else
                    strSql = strSql & "AND CharIndex('" & Session("Carteira") & "',FNEG_Carteiras) > 0"
                End If
                objDRMirror = Conn.ExecuteQuery(strSql)

                If objDRMirror.Read Then
                    'x = x + 1
                    '    'cboCondicao.Items.Add(Right("00" & Trim(objDRMirror("FNEG_QtdParcelas")), 2) & "x | " & Right("00" & Trim(objDRMirror("FNEG_DescPrincipal")), 2) & "%")
                    If CalcularParcelas(dblPrincipal, Session("valCorrecao"), x, Trim(objDRMirror("FNEG_DescPrincipal")), 100, Session("honor"), 0) >= 50 Then cboCondicao.Items.Add(Right("00" & Trim(objDRMirror("FNEG_QtdParcelas")), 2) & "x | Desc: R$" & FormatNumber(dblSaldo - CalcularParcelas(dblPrincipal, Session("valCorrecao"), 1, Trim(objDRMirror("FNEG_DescPrincipal")), 100, Session("honor"), 0), 2) & " (" & Right("00" & Trim(objDRMirror("FNEG_DescPrincipal")), 2) & "%)")
                    '    'arrFaixas(x) = CInt(Trim(objDR("FNEG_Id")))

                End If

                'cboCondicao.Items.Add(Right("00" & Trim(objDR("FNEG_QtdParcelas")), 2) & "x | " & Right("00" & Trim(objDR("FNEG_DescPrincipal")), 2) & "%")

            Loop

            Session("Saldo") = dblSaldo
            Session("Principal") = dblPrincipal
            'valCorrecao.Text = Session("valCorrecao")
            'Atualizado.Text = dblSaldo
            'Principal.Text = Session("Principal")
            'PercentHonor.Text = Session("PercentHonor")

            cboCondicao.SelectedIndex = 0
            lblSaldo.Text = "R$" & FormatNumber(CStr(dblSaldo), 2)
            'lblParcela.Text = "R$" & CStr(dblSaldo / CInt(Left(cboCondicao.SelectedItem.Value, InStr(cboCondicao.SelectedItem.Value, "/") - 1)))
            'lblParcela.Text = "R$" & CalculaSaldo(dblSaldo, arrFaixas(CInt(cboCondicao.SelectedIndex)))

        Catch ex As Exception
            SendMailErro("Erro na Negociação On-Line." & Chr(13) & Chr(13) & "Número: " & Err.Number & Chr(13) & "Source: " & Err.Source & Chr(13) & "Descrição: " & Err.Description & Chr(13) & "Help: " & Chr(13) & Err.HelpFile & Chr(13) & Err.HelpContext & Chr(13) & Err.LastDllError)
            ShowMessage("Sistema Inativo. Por favor tente novamente em alguns minutos ou entre em contato com nosso atendimento no telefone 11-2171-0301(SP) ou 0800-6000-500(Demais Localidades).")
            'Response.Redirect("inCPF.aspx")
        End Try

    End Sub

    Protected Sub ShowMessage(ByVal strMsg As String)

        lblMsg.Text = ""
        lblMsg.ForeColor = Drawing.Color.Red
        lblMsg.Visible = True
        lblMsg.Text = Chr(13) & strMsg & Chr(13)

    End Sub

    Protected Function Validar() As Boolean

        Dim datNow As String = Right("00" & Day(Now), 2) & "/" & Right("00" & Month(Now), 2) & "/" & Year(Now)
        Dim datVencto As String = Right("00" & Day(DateAdd(DateInterval.Day, 5, Now)), 2) & "/" & Right("00" & Month(DateAdd(DateInterval.Day, 5, Now)), 2) & "/" & Year(DateAdd(DateInterval.Day, 5, Now))
        Dim datLimite As String = Right("00" & Day(DateAdd(DateInterval.Day, 15, Now)), 2) & "/" & Right("00" & Month(DateAdd(DateInterval.Day, 15, Now)), 2) & "/" & Year(DateAdd(DateInterval.Day, 15, Now))

        lblMsg.ForeColor = Drawing.Color.DarkOrange
        lblMsg.Text = "Carregando..."
        lblMsg.Visible = False

        Try

            Validar = True

            If Trim(txtVencto.Text) = "" Then
                txtVencto.Text = datVencto
            End If

            If Not IsDate(Trim(txtVencto.Text)) Then
                ShowMessage("Vencimento da parcela deve estar no formato DD/MM/AAAA.")
                txtVencto.Focus()
                Validar = False
            End If

            If CInt(DatePart(DateInterval.Year, CDate(txtVencto.Text)) & Right("00" & DatePart(DateInterval.Month, CDate(txtVencto.Text)), 2) & Right("00" & DatePart(DateInterval.Day, CDate(txtVencto.Text)), 2)) < CInt(DatePart(DateInterval.Year, CDate(datVencto)) & Right("00" & DatePart(DateInterval.Month, CDate(datVencto)), 2) & Right("00" & DatePart(DateInterval.Day, CDate(datVencto)), 2)) Then
                ShowMessage("Vencimento da parcela deve ser maior que HOJE + 5 dias.<br>Caso deseje pagar antes dessa data, entre em contato com nosso atendimento no telefone 11-2171-0301(SP) ou 0800-6000-500(Demais Localidades).")
                txtVencto.Focus()
                Validar = False
            End If

            If CInt(DatePart(DateInterval.Year, CDate(txtVencto.Text)) & Right("00" & DatePart(DateInterval.Month, CDate(txtVencto.Text)), 2) & Right("00" & DatePart(DateInterval.Day, CDate(txtVencto.Text)), 2)) > CInt(DatePart(DateInterval.Year, CDate(datLimite)) & Right("00" & DatePart(DateInterval.Month, CDate(datLimite)), 2) & Right("00" & DatePart(DateInterval.Day, CDate(datLimite)), 2)) Then
                ShowMessage("Vencimento da parcela deve ser até " & datLimite & ".<br>Caso deseje pagar após essa data, entre em contato com nosso atendimento no telefone 11-2171-0301(SP) ou 0800-6000-500(Demais Localidades).")
                txtVencto.Focus()
                Validar = False
            End If

            If Trim(cboCondicao.SelectedItem.Text) = "Selecione..." Then
                ShowMessage("Selecione o parcelamento.")
                cboCondicao.Focus()
                Validar = False
            End If

            If Left(cboCondicao.SelectedItem.Text, 2) = "01" Then
                txtValor.Text = ""
            End If

            If dblEntradaMin <> 0 Then
                If Trim(txtValor.Text) <> "" And Trim(txtValor.Text) <> "0" And CDbl(Replace(Trim(txtValor.Text), "", "0")) < dblEntradaMin Then
                    ShowMessage("O valor mínimo da entrada para este parcelamento é R$" & Replace(dblEntradaMin, ".", ",") & ".<br>Para alterar o valor mínimo de entrada, mude a condição de parcelamento ou entre em contato com nosso atendimento no telefone 11-2171-0301(SP) ou 0800-6000-500(Demais Localidades).")
                    txtValor.Text = Replace(dblEntradaMin, ".", ",")
                    txtValor.Focus()
                    Validar = False
                End If
            End If

            If Trim(txtValor.Text) <> "" And Trim(txtValor.Text) <> "0" Then
                If CLng(Replace(txtValor.Text, ",", ".")) > CLng(Replace(Replace(lblSaldo.Text, ",", "."), "R$", "")) Then
                    ShowMessage("O valor da entrada deve ser menor que o valor do Acordo.")
                    txtValor.Focus()
                    Validar = False
                End If
            End If

        Catch ex As Exception
            Validar = False
        End Try

    End Function

    Protected Sub btnCalcularServer_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCalcularServer.Click

        Try

            'lblParcela.Text = "Calculando..."
            cboCondicao.Enabled = False
            txtVencto.Enabled = False
            txtValor.Enabled = False
            btnCalcularServer.Enabled = False
            btnNegociacao.Enabled = False

            If Validar() = False Then Exit Try

            If Left(cboCondicao.SelectedItem.Text, 2) = "01" Then txtValor.Text = 0

            'function validaCampos(){
            '	    var objDateNow = new Date();
            '	    if(document.form1.txtVencto.value==""){
            '	        var objDateVencto = new Date();
            '		    var objDateVencto = objDateVencto.add("d",5);
            '			document.form1.txtVencto.value = Right("00" + parseInt(objDateVencto.getDate()),2) + "/" + Right("00" + parseInt(parseInt(objDateVencto.getMonth()) + parseInt(1)),2) + "/" + objDateVencto.getFullYear();
            '		}
            '		if(! fnIsDate(document.form1.txtVencto.value)){
            '			alert("Vencimento da parcela deve estar no formato DD/MM/AAAA.");
            '			document.form1.txtVencto.select();
            '			return false;
            '		}

            '		var objDate = new Date();
            '		var objDateAdd = new Date();
            '		var objDateAdd = objDateAdd.add("d",15);
            '		var datToday = (objDateAdd.getFullYear() + "" + Right("00" + parseInt(parseInt(objDateAdd.getMonth()) + parseInt(1)),2) + "" + Right("00" + parseInt(objDateAdd.getDate()),2));
            '		var datNow = (objDate.getFullYear() + "" + Right("00" + parseInt(parseInt(objDate.getMonth()) + parseInt(1)),2) + "" + Right("00" + parseInt(objDate.getDate()),2));
            '		var arrDat_Prioridade = document.form1.txtVencto.value.split("/");
            '		var datDat_Prioridade = (arrDat_Prioridade[2] + "" + Right("00" + parseInt(arrDat_Prioridade[1]),2) + "" + Right("00" + parseInt(arrDat_Prioridade[0]),2));

            '//        alert(datDat_Prioridade + " " + datToday + " " + datNow);
            '        if(parseInt(datDat_Prioridade) <= parseInt(datNow)){
            '			alert("Vencimento da parcela deve ser maior que hoje.");
            '			document.form1.txtVencto.select();
            '			return false;
            '		}

            '		if(parseInt(datDat_Prioridade) > parseInt(datToday)){
            '			alert("O vencimento da parcela deve ser menor que hoje + 15 dias.");
            '			document.form1.txtVencto.select();
            '			return false;
            '		}

            '		if(parseInt(document.form1.txtValor.value) < 50){
            '		    alert("O valor mínimo da entrada é de R$50,00.");
            '			document.form1.txtValor.select();
            '			return false;
            '		}

            '		//alert(parseInt(Replace(document.form1.txtValor.value,".","")) + " " + parseInt(document.form1.Atualizado.value))
            '		if(parseInt(Replace(document.form1.txtValor.value,".","")) > parseInt(document.form1.Atualizado.value)){
            '			alert("O valor da entrada deve ser menor que o valor do Acordo.");
            '			document.form1.txtVencto.select();
            '			return false;
            '		}

            '		if(document.form1.txtValor.value==""){
            '			document.form1.txtValor.value='0';
            '		}

            '		if(Left(document.form1.cboCondicao.value,2) == "01"){
            '			document.form1.txtValor.value='0';
            '		}

            '		if(document.form1.cboCondicao.value=="Selecione..."){
            '			alert("Selecione o parcelamento.");
            '			document.form1.cboCondicao.select();
            '			return false;
            '		}
            '		return true;
            '	}

            If dblEntradaMin = 0 And Trim(txtValor.Text) <> "" Then txtValor.Text = ""

            lblParcela.Text = Calcular( _
                              Trim(Session("Principal")), _
                              Trim(Session("valCorrecao")), _
                              Trim(Session("PercentHonor")), _
                              Trim(Replace(cboCondicao.SelectedItem.Text, "x", "")), _
                              Replace(txtValor.Text, ".", ""), _
                              Trim(txtVencto.Text), _
                              Trim(idLogin), _
                              Trim(strContrato))

            If Trim(lblParcela.Text) = "msg" Then
                lblParcela.Text = ""
            Else
                lblMsg.Text = ""
                lblMsg.Visible = False
            End If

        Catch ex As Exception
            SendMailErro("Erro na Negociação On-Line." & Chr(13) & Chr(13) & "Número: " & Err.Number & Chr(13) & "Source: " & Err.Source & Chr(13) & "Descrição: " & Err.Description & Chr(13) & "Help: " & Chr(13) & Err.HelpFile & Chr(13) & Err.HelpContext & Chr(13) & Err.LastDllError)
            ShowMessage("Sistema Inativo. Por favor tente novamente em alguns minutos ou entre em contato com nosso atendimento no telefone 11-2171-0301(SP) ou 0800-6000-500(Demais Localidades).")
        End Try

        If lblParcela.Text = "Calculando..." Then lblParcela.Text = ""
        cboCondicao.Enabled = True
        txtVencto.Enabled = True
        txtValor.Enabled = True
        btnCalcularServer.Enabled = True
        btnNegociacao.Enabled = True

    End Sub

    '<Ajax.AjaxMethod()> _
    Public Function Calcular(ByVal valPrincipal As Double, ByVal valCorrecao As Double, ByVal percentHonor As Double, ByVal Condicao As String, ByVal valEntrada As Double, ByVal datEntrada As String, ByVal lngLoginId As Long, ByVal strContratoToOS As String) As String

        Try

            Dim strRetorno As String = ""
            Dim arrCondicao() As String
            Dim qtdParcelas As Integer = 0
            Dim percentDescPrinc As Double = 0
            Dim percentDescCorr As Double = 0
            Dim ValorParcela As Double = 0
            Dim strEntrada As String = ""
            Dim strParcela As String = ""
            Dim x As Integer = 0
            Dim strValor As String = ""
            Dim strTotal As String = ""
            Dim descNeg As String = ""
            Dim strSql As String = ""
            Dim Mirror As New Comando
            Dim strDataEntrada As String = ""
            Dim parcToOS As Integer = 0
            Dim y As Integer = 0
            Dim z As Integer = 0
            Dim totalComDesc As Double = 0

            Mirror.Banco = "MIRRORWEB"

            'If percentHonor = 20 Then percentHonor = 19.99

            arrCondicao = Split(Trim(Condicao), "|")
            qtdParcelas = arrCondicao(0)
            percentDescPrinc = Mid(arrCondicao(1), InStr(arrCondicao(1), "%") - 2, 2)
            percentDescCorr = 100

            If idLogin = 0 Then idLogin = lngLoginId
            If Trim(strContrato) = "" Then strContrato = Trim(strContratoToOS)

            ValorParcela = CalcularParcelas(valPrincipal, valCorrecao, qtdParcelas, percentDescPrinc, percentDescCorr, percentHonor, valEntrada)

            If ValorParcela = 0 Then
                'SendMailErro("Erro na rotina Calcular.")
                'ShowMessage("Sistema Inativo. Por favor tente novamente em alguns minutos.")
                'ShowMessage("Negociação Indisponível - Tente mais tarde.")
                'Response.Redirect("inCPF.aspx")
                Return "msg"
                Exit Try
            End If

            If ValorParcela = 88888888 Then
                'Return "Sistema Inativo. Por favor tente novamente em alguns minutos."
                Exit Try
            End If

            If ValorParcela < 50 Then
                totalComDesc = CalcularParcelas(valPrincipal, valCorrecao, 2, percentDescPrinc, percentDescCorr, percentHonor, valEntrada)
                If InStr(totalComDesc / 50, ",") > 0 Then
                    y = Left(totalComDesc / 50, InStr(totalComDesc / 50, ",") - 1)
                Else
                    y = totalComDesc / 50
                End If
                ShowMessage("Para entrada de R$" & FormatNumber(Replace(valEntrada, ".", ","), 2) & " o parcelamento máximo é " & y + 1 & "x.")
                Return "msg"
                Exit Try
            End If

            dblEntradaMin = ValorParcela
            Session("EntradaMin") = dblEntradaMin
            Session("Parc") = qtdParcelas

            descNeg = ""
            strDataEntrada = Format(CDate(datEntrada), "yyy-MM-dd")

            For x = 1 To qtdParcelas

                If valEntrada <> 0 And x = 1 Then
                    strValor = Replace(FormatNumber(valEntrada, 2), ".", ",")
                    descNeg = "<b>Contrato: </b>" & strContrato & "<br><b>Entrada: </b>R$" & strValor & "<br>"
                    strTotal = Replace(FormatNumber(valEntrada + ((qtdParcelas - 1) * ValorParcela), 2), ".", ",")
                ElseIf x = 1 Or (x = 2 And valEntrada <> 0) Then
                    strValor = Replace(FormatNumber(ValorParcela, 2), ".", ",")
                    If Trim(strTotal) = "" Then strTotal = Replace(FormatNumber(qtdParcelas * ValorParcela, 2), ".", ",")
                End If

                If x > 1 Then
                    datEntrada = Format(DateAdd(DateInterval.Month, 1, CDate(datEntrada)), "dd/MM/yyy")
                Else
                    strRetorno = "<table width=""180"" border=""0"" cellspacing=""3"" cellpadding=""0""><tr align=""center""><td width=""35""><b>N&ordm;</b></td><td width=""80""><b>Data</b></td><td align=""right"" width=""65""><b>Valor</b></td></tr>"
                End If

                strRetorno = strRetorno & "<tr><td align=""center"">" & Right("00" & x, 2) & "</td><td align=""center"">" & datEntrada & "</td><td align=""right"">" & strValor & "</td></tr>"

            Next x

            If valEntrada = 0 Then parcToOS = qtdParcelas Else parcToOS = qtdParcelas - 1

            strRetorno = strRetorno & "<tr><td align=""center""></td><td align=""right"">Total</td><td align=""right"">" & strTotal & "</td></tr></table>"
            descNeg = descNeg & "<b>Parcelas: </b> "
            If valEntrada <> 0 Then descNeg = descNeg & "+ "
            descNeg = descNeg & parcToOS & "x de R$" & Replace(FormatNumber(ValorParcela, 2), ".", ",") & "<br>"
            descNeg = descNeg & "<b>Data 1º Vencto: </b>" & Format(CDate(strDataEntrada), "dd/MM/yyy") & "<br>"
            descNeg = descNeg & "<b>Total do Acordo: </b>R$" & Replace(FormatNumber(valEntrada + (parcToOS * ValorParcela), 2), ".", ",")

            'Session("descNeg") = Trim(descNeg)

            strSql = "prc_INS_NegSimulacao " & idLogin & "," & Replace(Round(valPrincipal + valCorrecao + (percentHonor * ((valPrincipal + valCorrecao) / 100)), 2), ",", ".") & "," & Replace(valPrincipal, ",", ".") & "," & qtdParcelas & "," & Replace(percentDescPrinc, ",", ".") & "," & Replace(percentDescCorr, ",", ".") & "," & Replace(percentHonor, ",", ".") & "," & Replace(Trim(CStr(valEntrada)), ",", ".") & "," & Replace(ValorParcela, ",", ".") & ",'" & strDataEntrada & "','" & descNeg & "'"
            Mirror.Execute(strSql)

            If Trim(txtValor.Text) <> "" Then
                If CDbl(txtValor.Text) < dblEntradaMin Then
                    txtValor.Text = Replace(dblEntradaMin, ".", ",")
                End If
            Else
                txtValor.Text = Replace(dblEntradaMin, ".", ",")
            End If
            txtValor.Enabled = True

            Return strRetorno

        Catch ex As Exception

            SendMailErro("Erro na Negociação On-Line." & Chr(13) & Chr(13) & "Número: " & Err.Number & Chr(13) & "Source: " & Err.Source & Chr(13) & "Descrição: " & Err.Description & Chr(13) & "Help: " & Chr(13) & Err.HelpFile & Chr(13) & Err.HelpContext & Chr(13) & Err.LastDllError)
            ShowMessage("Sistema Inativo. Por favor tente novamente em alguns minutos ou entre em contato com nosso atendimento no telefone 11-2171-0301(SP) ou 0800-6000-500(Demais Localidades).")
            'ShowMessage("Negociação Indisponível - Tente mais tarde.")
            'Response.Redirect("inCPF.aspx")
            Return "msg"

        End Try

    End Function

    Protected Function CalcularParcelas(ByVal valPrincipal As Double, ByVal valCorrecao As Double, ByVal qtdParcelas As Integer, ByVal percentDescPrinc As Double, ByVal percentDescCorr As Double, ByVal percentHonor As Double, ByVal valEntrada As Double) As Double

        Dim descPrincipal As Double = 0
        Dim descCorrecao As Double = 0
        Dim valHonorarios As Double = 0

        Try

            descPrincipal = FormatNumber((percentDescPrinc * (valPrincipal / 100)), 2)
            descCorrecao = FormatNumber((percentDescCorr * (valCorrecao / 100)), 2)
            valHonorarios = FormatNumber((percentHonor * ((valPrincipal - descPrincipal) / 100)), 2) '+ 0.4

            'descPrincipal = Round((percentDescPrinc * (valPrincipal / 100)), 2)
            'descCorrecao = Round((percentDescCorr * (valCorrecao / 100)), 2)
            'valHonorarios = Round((percentHonor * ((valPrincipal - descPrincipal) / 100)), 2) '+ 0.4

            If (valPrincipal - descPrincipal + valHonorarios + valCorrecao - descCorrecao - 12) < valEntrada Then
                ShowMessage("Valor da entrada é maior que o total da dívida com os descontos. <br>Insira um valor menor que R$" & Replace((valPrincipal - descPrincipal + valHonorarios + valCorrecao - descCorrecao - 12), ".", ",") & ".")
                Exit Try
            End If

            If valEntrada = 0 Then
                CalcularParcelas = FormatNumber((valPrincipal - descPrincipal + valHonorarios + valCorrecao - descCorrecao) / qtdParcelas, 2)
                'CalcularParcelas = Round((valPrincipal - descPrincipal + valHonorarios + valCorrecao - descCorrecao) / qtdParcelas, 2)
            Else
                CalcularParcelas = FormatNumber((valPrincipal - descPrincipal + valHonorarios + valCorrecao - descCorrecao - valEntrada) / (qtdParcelas - 1), 2)
                'CalcularParcelas = Round((valPrincipal - descPrincipal + valHonorarios + valCorrecao - descCorrecao - valEntrada) / (qtdParcelas - 1), 2)
            End If

        Catch ex As Exception
            SendMailErro("Erro na Negociação On-Line." & Chr(13) & Chr(13) & "Número: " & Err.Number & Chr(13) & "Source: " & Err.Source & Chr(13) & "Descrição: " & Err.Description & Chr(13) & "Help: " & Chr(13) & Err.HelpFile & Chr(13) & Err.HelpContext & Chr(13) & Err.LastDllError)
            CalcularParcelas = 88888888
            'ShowMessage("Negociação Indisponível - Tente mais tarde.")
            ShowMessage("Negociação Indisponível - Tente mais tarde ou entre em contato com nosso atendimento no telefone 11-2171-0301(SP) ou 0800-6000-500(Demais Localidades).")
            'Response.Redirect("inCPF.aspx")
        End Try

    End Function

    Protected Sub btnNegociacao_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNegociacao.Click

        Dim strSql As String = ""
        Dim objDR As SqlDataReader = Nothing
        Dim Conn As New Comando
        Dim descNeg As String = ""
        Dim simID As Integer = 0

        Try

            btnNegociacao.Enabled = False
            btnCalcularServer.Enabled = False

            If Validar() = False Then
                btnNegociacao.Enabled = True
                btnCalcularServer.Enabled = True
                Exit Sub
            End If

            lblParcela.Text = Calcular( _
                              Trim(Session("Principal")), _
                              Trim(Session("valCorrecao")), _
                              Trim(Session("PercentHonor")), _
                              Trim(Replace(cboCondicao.SelectedItem.Text, "x", "")), _
                              Replace(txtValor.Text, ".", ""), _
                              Trim(txtVencto.Text), _
                              Trim(idLogin), _
                              Trim(strContrato))

            Conn.Banco = "MIRRORWEB"
            strSql = "prc_SEL_NegNegociacao " & idLogin
            objDR = Conn.ExecuteQuery(strSql)

            If Not objDR.Read Then
                ShowMessage("Clique em calcular antes de registrar uma negociação. É necessário um cálculo válido.")
                Exit Try
            End If

            descNeg = Trim(objDR("DescNeg").ToString)
            Session("descNeg") = Trim(descNeg)
            simID = Trim(objDR("SimID").ToString)
            Session("SimID") = Trim(simID)

            GravarInput(strContrato, "<b>Simulação de Negociação On-Line<br>Descrição:<br>" & descNeg & "<br>Data: " & Format(CDate(Now), "dd/MM/yyy HH:mm:ss") & "</b>", "N_ONLINE")

            Response.Redirect("inConfirmacao.aspx?id=" & idLogin & "&idcli=" & idCliente & "&ctra=" & strContrato & "&simid=" & simID, False)

        Catch ex As Exception
            btnNegociacao.Enabled = True
            btnCalcularServer.Enabled = True
            SendMailErro("Erro na Negociação On-Line." & Chr(13) & Chr(13) & "Número: " & Err.Number & Chr(13) & "Source: " & Err.Source & Chr(13) & "Descrição: " & Err.Description & Chr(13) & "Help: " & Chr(13) & Err.HelpFile & Chr(13) & Err.HelpContext & Chr(13) & Err.LastDllError)
            ShowMessage("Sistema Inativo. Por favor tente novamente em alguns minutos ou entre em contato com nosso atendimento no telefone 11-2171-0301(SP) ou 0800-6000-500(Demais Localidades).")
            'Response.Redirect("inCPF.aspx")
        End Try

    End Sub

    Protected Function AtualizaValor(ByVal ID_Carteira As Integer, ByVal VencimentoDebito As String, ByVal AtualizaAte As String, ByVal Valor As Double, ByVal Sinal As String, ByVal PermiteAtualizacao As Boolean, ByVal QtdParc As Integer, ByVal DataRecebimento As String, ByVal TDOC_ID As Integer, ByVal DataAtualizacao As String, ByVal DataVencDebito As String)
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

        Try

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
                vPercentHonorarios = Trim(Session("honor"))
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
                    If Trim(Session("PercentHonor")) = "" Then Session("PercentHonor") = vPercentHonorarios
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
                            If Trim(Session("valCorrecao")) <> "" Then
                                Session("valCorrecao") = Round(CDbl(Session("valCorrecao")) + Round((DRTaxaHoje("COTA_Indice") / DRTaxaVenc("COTA_Indice")) * vValorCorrigido), 2)
                            Else
                                Session("valCorrecao") = Round((DRTaxaHoje("COTA_Indice") / DRTaxaVenc("COTA_Indice")) * vValorCorrigido)
                            End If
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
                                        'If Trim(Session("valCorrecao")) = "" Then Session("valCorrecao") = Round((DRIGPMHoje("COTA_Indice") / DRIGPMVencimento("COTA_Indice")) * vValorCorrigido)
                                        If Trim(Session("valCorrecao")) <> "" Then
                                            Session("valCorrecao") = Round(CDbl(Session("valCorrecao")) + Round((DRIGPMHoje("COTA_Indice") / DRIGPMVencimento("COTA_Indice")) * vValorCorrigido), 2)
                                        Else
                                            Session("valCorrecao") = Round(Round((DRIGPMHoje("COTA_Indice") / DRIGPMVencimento("COTA_Indice")) * vValorCorrigido), 2)
                                        End If
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
                                        If Trim(Session("valCorrecao")) = "" Then Session("valCorrecao") = Round(((DRIGPMHoje("COTA_Indice") / DRIGPMVencimento("COTA_Indice")) * vValorCorrigido) - vValorCorrigido, 2)
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
                                If Trim(Session("valCorrecao")) <> "" Then
                                    Session("valCorrecao") = Round(CDbl(Session("valCorrecao")) + vValorJuros, 2)
                                Else
                                    Session("valCorrecao") = Round(vValorJuros, 2)
                                End If
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

        Catch ex As Exception
            SendMailErro("Erro na Negociação On-Line." & Chr(13) & Chr(13) & "Número: " & Err.Number & Chr(13) & "Source: " & Err.Source & Chr(13) & "Descrição: " & Err.Description & Chr(13) & "Help: " & Chr(13) & Err.HelpFile & Chr(13) & Err.HelpContext & Chr(13) & Err.LastDllError)
            ShowMessage("Negociação Indisponível - Tente mais tarde ou entre em contato com nosso atendimento no telefone 11-2171-0301(SP) ou 0800-6000-500(Demais Localidades).")
            'Response.Redirect("inCPF.aspx")
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
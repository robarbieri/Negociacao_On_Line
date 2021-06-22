<%
'Função para descrever informações sobre a parcela
'strContrato: Contrato
'strContratante: Nome do Contratante
'strParcPlano: Parcela / Plano
'dtVenc: Vencimento
'strTel: Telefone
'dblValParc: Valor da parcela
sub DescBoletoParcela(strContrato, strContratante, strParcPlano, dtVenc, strTel, dblValParc)
%>
	<TABLE WIDTH="640" CELLSPACING=0 CELLPADDING=0 BORDER=0 ID="Table1">
		<tr><td><img src="../images/logo_cliente_logon_26.gif">	</td></tr>
	</TABLE>
	<TABLE WIDTH="640" CELLSPACING=0 CELLPADDING=0 BORDER=0 ID="Table2" class="tabela" height="420">
		<tr>
			<td>	
					<table width="90%" align=center>
					<tr>
						<td>&nbsp;&nbsp; </td>
					</tr>
		
					<tr>
						<td class="texto1">Prezado(a) Senhor(a), </td>
					</tr>

					<tr>
						<td>&nbsp;&nbsp; </td>
					</tr>
		
					<tr>
						<%if vCONT_ID = 53 or vCONT_ID = 101 or vCONT_ID = 65 then%>
							<td class="texto1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Este boleto refere-se à parcela do acordo firmado por V.Sa., para pagamento do saldo remanescente do contrato: <%=FormataContrato(strContrato)%> da <b><%=UCase(strContratante)%></b>.</td>
						<%else%>
							<td class="texto1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Esta fatura refere-se à parcela do acordo firmado por V.Sa., para pagamento de seu débito junto a <b><%=UCase(strContratante)%>:</b> <%=FormataContrato(strContrato)%> </td>
						<%end if%>
					</tr>
					
					<tr>
						<td>&nbsp;&nbsp; </td>
					</tr>

					<tr>
						<td>
								<TABLE WIDTH="70%" align="center" BORDER=1 CELLSPACING=0 CELLPADDING=1 ID="Table11">
										<tr>
											<td class="texto1" align="center">Parcela / Plano</td>
											<td class="texto1" align="center">Valor da Parcela</td>
											<td class="texto1" align="center">Vencimento</td>
										</tr>
										<tr>
											<td class="texto1" align="center"><%=strParcPlano%></td>
											<td class="texto1" align="center">R$ <%=FormatNumber(dblValParc)%></td>
											<td class="texto1" align="center"><%=dtVenc%></td>
										</tr>
								</table>
						</td>
					</tr>

					<tr>
						<td>&nbsp;&nbsp; </td>
					</tr>
					
					<tr>
						<td class="texto1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; A pontualidade dos seus pagamentos é muito importante para futura avaliação de crédito. </td>
					</tr>
					<tr>
						<td>&nbsp;&nbsp; </td>
					</tr>

					<tr>
						<td class="texto1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Se você receber esta fatura após a data limite para pagamento, entre em contato com nossa Central de Atendimento ao Cliente. </td>
					</tr>
					<tr>
						<td>&nbsp;&nbsp; </td>
					</tr>

					<tr>
						<td align=center class="texto1"> <b>... Dúvidas ... </b> </td>
					</tr>
					<tr>
						<td align=center class="texto1"> <b>Central de Atendimento ao Cliente</b> </td>
					</tr>
					<tr>
						<td align=center class="texto1"> <b><%=strTel%></b> </td>
					</tr>
					
					<tr>
						<td>&nbsp;&nbsp; </td>
					</tr>
					
					<tr>
						<td align="center" class="texto1"> Caso esta parcela já esteja paga desconsidere esta correspondência.</td>
					</tr>
					<tr>
						<td>&nbsp;&nbsp; </td>
					</tr>
				
					</table>
			</td>
		</tr>
	</TABLE>
<%
end sub
%>

<%
'função que mostra informações sobre o boleto dos clientes da empresario
'idcontratante = codigo do contratante (CONT_ID)
sub DescBoletoEmpresario(idcontratante)
%>
	<TABLE WIDTH="640" CELLSPACING=0 CELLPADDING=0 BORDER=0 ID="Table31">
  		<TR>
			<TD class=cp VALIGN=BOTTOM WIDTH=225><IMG SRC="../images/logo_cliente_logon_25.gif"></TD>
			<%
			if vCONT_ID = 48 then
			%>
			<TD ALIGN=RIGHT VALIGN=BOTTOM><IMG SRC="../images/logo_cea.JPG"></TD>
			<%
			else
			%>
			<TD ALIGN=RIGHT VALIGN=BOTTOM>&nbsp;</TD>
			<%
			end if
			%>
		</TR>

  		<TR>
			<TD colspan=2><br>
			
				<TABLE WIDTH="460" ALIGN=RIGHT CELLSPACING=0 CELLPADDING=0 BORDER=0 ID="Table4">
						<TR>
							<TD valign=top align=left>
								<FONT class=cn><b>
									<%=var_sacado%><BR>
									<%=var_endereco%>,&nbsp;<%=var_bairro%><br>
									<%=var_cidade%>&nbsp;-&nbsp;<%=var_estado%>&nbsp;-&nbsp;CEP: <%=var_cep%><BR></b>
								</FONT>
	 						</TD>
						</TR>
					</TABLE>
				<br><br><br>
				<TABLE id="Table43" borderColor=black cellSpacing=0 borderColorDark=white cellPadding=0 width=640 border=1>
					<TBODY>
					<TR>
					<TD width="12%" bgColor=#d3d3d3><FONT style="FONT-SIZE: 8pt" face=Arial>&nbsp;<B>CONTRATO</FONT></B></TD>
					<TD style="FONT-SIZE: 8pt" width="38%"><FONT style="FONT-SIZE: 8pt" FACE="Arial">&nbsp;<% = FormataContrato(rsParcelas("CTRA_Numero"))%></FONT></TD>
					<TD width="12%" bgColor=#d3d3d3><FONT style="FONT-SIZE: 8pt" face=Arial>&nbsp;<B>CPF</FONT></B></TD>
					<TD style="FONT-SIZE: 8pt" width="38%"><FONT style="FONT-SIZE: 8pt" FACE="Arial">&nbsp;<% = FormataCPFCNPJ(rsParcelas("DEVE_CGCCPF"))%></FONT></TD>
					</TR>
					<TR>
					<TD bgColor=#d3d3d3><FONT style="FONT-SIZE: 8pt" face=Arial>&nbsp;<B>CREDOR</FONT></B></TD>
					<TD><FONT style="FONT-SIZE: 8pt" face=Arial>&nbsp;<% = vNomeEmpresa%></FONT></TD>
					<TD bgColor=#d3d3d3><FONT style="FONT-SIZE: 8pt" face=Arial>&nbsp;<B>CÓDIGO</FONT></B></TD>
					<TD><FONT style="FONT-SIZE: 8pt" face=Arial>&nbsp;</FONT></TD>
					</TR>
					</TBODY>
				</TABLE>
				<TABLE width=640 border=0 ID="Table5">
				<TBODY>
					<TR>
						<TD borderColor=black borderColorDark=white>
							<TABLE borderColor=black cellSpacing=0 borderColorDark=white cellPadding=0 width="100%" border=1 ID="Table6">
							<TBODY>
								<TR bgColor=#d3d3d3>
									<TD align=middle width="100%" colSpan=5><FONT style="FONT-SIZE: 10pt" face=Arial><B>DEMONSTRATIVO DOS VALORES DO CONTRATO ORIGINAL EM ABERTO</B></FONT></TD>
								</TR>
								<TR style="FONT-WEIGHT: normal; FONT-SIZE: 8pt; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal" align=middle bgColor=#d3d3d3><FONT face=Arial>
									<TD>Parcela</TD>
									<TD>Vencimento</TD>
									<TD>Original</TD>
									<TD>Negociado</TD>
									<TD>Total</TD>
								</TR>
								<%
								
								vValorPrincipalAtual = 0
								vValorCorrigido = 0
								Set rsTit = Bd.Execute("SELECT * FROM Transacoes_de_Acordo t WITH (NOLOCK) LEFT JOIN Tipos_de_Documento td WITH (NOLOCK) ON t.TDOC_ID = td.TDOC_ID WHERE ACOR_ID = " & rsParcelas("ACOR_ID"))
								if rsTit.EOF then
									Set rsTit = Bd.Execute("SELECT * FROM Transacoes t WITH (NOLOCK) LEFT JOIN Tipos_de_Documento td WITH (NOLOCK) ON t.TDOC_ID = td.TDOC_ID WHERE CTRA_ID = " & rsParcelas("CTRA_ID"))
								end if
								Do While not rsTit.EOF
									if rsTit("TDOC_ID") = "" or isnull(rsTit("TDOC_ID")) then
										vValorPrincipalAtual = vValorPrincipalAtual + rsTit("TRAN_Valor")
										if rsTit("TRAN_Vencimento") > rsParcelas("ACOR_Data") then
											vValorCorrigido = vValorCorrigido
										else
											vValorCorrigido = vValorCorrigido + AtualizaValor(rsParcelas("CART_ID"), rsTit("TRAN_Vencimento"), rsParcelas("ACOR_Data"), rsTit("TRAN_Valor"), "+", 1, rsParcelas("Plano"), rsParcelas("CTRA_DataRecebimentoContrato"), -1, rsTit("TRAN_DataRecebimento"), rsTit("TRAN_Vencimento"))
										end if
									else
										if rsTit("TRAN_Vencimento") > rsParcelas("ACOR_Data") then
											vValorCorrigido = vValorCorrigido
										else
											vValorCorrigido = eval(CStr(Replace(Replace(vValorCorrigido,".",""),",",".")) & CStr(rsTit("TDOC_Sinal")) & CStr(Replace(Replace(AtualizaValor(rsParcelas("CART_ID"), rsTit("TRAN_Vencimento"), rsParcelas("ACOR_Data"), rsTit("TRAN_Valor"), rsTit("TDOC_Sinal"), 1, rsParcelas("Plano"), rsParcelas("CTRA_DataRecebimentoContrato"), -1, rsTit("TRAN_DataRecebimento"), rsTit("TRAN_Vencimento")),".",""),",",".")))
										end if
										vValorPrincipalAtual = eval(CStr(Replace(Replace(vValorPrincipalAtual,".",""),",",".")) & CStr(rsTit("TDOC_Sinal")) & CStr(Replace(Replace(rsTit("TRAN_Valor"),".",""),",",".")))
									end if
									rsTit.MoveNext
								loop
																
								vPercentDesc = 1 - (rsParcelas("ACOR_ValorAcordo") / vValorCorrigido)
								Set rsTit = Bd.Execute("SELECT * FROM Transacoes_de_Acordo t WITH (NOLOCK) LEFT JOIN Tipos_de_Documento td WITH (NOLOCK) ON t.TDOC_ID = td.TDOC_ID WHERE ACOR_ID = " & rsParcelas("ACOR_ID"))
								if rsTit.EOF then
									Set rsTit = Bd.Execute("SELECT * FROM Transacoes t WITH (NOLOCK) LEFT JOIN Tipos_de_Documento td WITH (NOLOCK) ON t.TDOC_ID = td.TDOC_ID WHERE CTRA_ID = " & rsParcelas("CTRA_ID"))
								end if
								vTotalOriginal = 0
								vTotalNegociado = 0
								Do while not rsTit.EOF
									if isnull(rsTit("TDOC_Sinal")) then
										vValAt = AtualizaValor(rsParcelas("CART_ID"), rsTit("TRAN_Vencimento"), rsParcelas("ACOR_Data"), rsTit("TRAN_Valor"), "+", 1, rsParcelas("Plano"), rsParcelas("CTRA_DataRecebimentoContrato"), -1, rsTit("TRAN_DataRecebimento"), rsTit("TRAN_Vencimento"))
									else
										vValAt = AtualizaValor(rsParcelas("CART_ID"), rsTit("TRAN_Vencimento"), rsParcelas("ACOR_Data"), rsTit("TRAN_Valor"), rsTit("TDOC_Sinal"), 1, rsParcelas("Plano"), rsParcelas("CTRA_DataRecebimentoContrato"), -1, rsTit("TRAN_DataRecebimento"), rsTit("TRAN_Vencimento"))
									end if
									vValorOriginal = vValAt
									vValDesc = Round(vValAt * vPercentDesc, 2)
									vValorLiquido = Round(vValAt - vValDesc, 2)
									%>
									<TR>
										<TD align=middle><FONT style="FONT-SIZE: 8pt" face=Arial><% = rsTit("TRAN_NumTitulo")%></FONT></TD>
										<TD align=middle><FONT style="FONT-SIZE: 8pt" face=Arial><% = rsTit("TRAN_Vencimento")%></FONT></TD>
										<TD align=right><FONT style="FONT-SIZE: 8pt" face=Arial><% = rsTit("TDOC_Sinal") & FormatNumber(vValAt)%></FONT></TD>
										<TD align=right><FONT style="FONT-SIZE: 8pt" face=Arial><% = rsTit("TDOC_Sinal") & FormatNumber(vValorLiquido)%></FONT></TD>
										<TD align=right><FONT style="FONT-SIZE: 8pt" face=Arial><% = rsTit("TDOC_Sinal") & FormatNumber(vValorLiquido)%></FONT></TD>
									</TR>
									<%
									if rsTit("TDOC_Sinal") = "-" then
										vTotalOriginal = vTotalOriginal - vValorOriginal
										vTotalNegociado = vTotalNegociado - vValorLiquido
									else
										vTotalOriginal = vTotalOriginal + vValorOriginal
										vTotalNegociado = vTotalNegociado + vValorLiquido
									end if
									rsTit.MoveNext
								Loop

								%>
								<TR bgColor=#d3d3d3>
									<TD align=middle colSpan=2><FONT style="FONT-SIZE: 8pt" face=Arial><B>TOTAL</FONT></B></TD>
									<TD align=right><FONT style="FONT-SIZE: 8pt" face=Arial><B><% = vTotalOriginal%></FONT></B></TD>
									<TD align=right><FONT style="FONT-SIZE: 8pt" face=Arial><B><% = vTotalNegociado%></FONT></B></TD>
									<TD align=right><FONT style="FONT-SIZE: 8pt" face=Arial><B><% = vTotalNegociado%></FONT></B></TD>
								</TR>
							</TBODY>
							</TABLE>
						</TD>
						<TD>
							<TABLE borderColor=black cellSpacing=0 borderColorDark=white cellPadding=0 width="100%" border=1 ID="Table7">
							<TBODY>
								<TR bgColor=#d3d3d3>
									<TD align=middle width="100%" colSpan=7><FONT style="FONT-SIZE: 10pt" face=Arial><B>DEMONSTRATIVO DAS PARCELAS DO ACORDO FIRMADO EM <% = rsParcelas("ACOR_Data")%></B></FONT></TD>
								</TR>
								<TR bgColor=#d3d3d3>
									<TD align=middle><FONT style="FONT-SIZE: 8pt" face=Arial> Nº </FONT></TD>
									<TD align=middle><FONT style="FONT-SIZE: 8pt" face=Arial> Vencimento </FONT></TD>
									<TD align=middle><FONT style="FONT-SIZE: 8pt" face=Arial> Parcela </FONT></TD>
									<TD align=middle><FONT style="FONT-SIZE: 8pt" face=Arial>IOF + Boleto</FONT></TD>
									<TD align=middle><FONT style="FONT-SIZE: 8pt" face=Arial> Total </FONT></TD>
									<TD align=middle><FONT style="FONT-SIZE: 8pt" face=Arial> Status </FONT></TD>
									<TD align=middle><FONT style="FONT-SIZE: 8pt" face=Arial> Saldo </FONT></TD>
								</TR>
								<%
								vTotalParcelas = 0
								vTotalBoleto = 0
								vTotalSaldo = 0
								Set rsParc = BD.Execute("SELECT * FROM Parcelas WITH (NOLOCK) WHERE ACOR_ID = " & rsParcelas("ACOR_ID") & " ORDER BY PARC_Numero")
								Do While not rsParc.EOF
									if IsNull(rsParc("PARC_DataPagamento")) then 
								%>
								<TR>
									<TD align=middle><FONT style="FONT-SIZE: 8pt" face=Arial><% = rsParc("PARC_Numero")%> </FONT></TD>
									<TD align=middle><FONT style="FONT-SIZE: 8pt" face=Arial><% = rsParc("PARC_Vencimento")%> </FONT></TD>
									<TD align=right><FONT style="FONT-SIZE: 8pt" face=Arial><% = FormatNumber(rsParc("PARC_ValorTotal")) %> </FONT></TD>
									<TD align=right><FONT style="FONT-SIZE: 8pt" face=Arial><% = FormatNumber(vValorTaxaBoleto)%></FONT></TD>
									<TD align=right><FONT style="FONT-SIZE: 8pt" face=Arial><% = FormatNumber(rsParc("PARC_ValorTotal") + vValorTaxaBoleto) %> </FONT></TD>
									<TD align=middle><FONT style="FONT-SIZE: 8pt" face=Arial>Aberto</FONT></TD>
									<TD align=right><FONT style="FONT-SIZE: 8pt" face=Arial><% = FormatNumber(rsParc("PARC_ValorTotal") + vValorTaxaBoleto) %> </FONT></TD>
								</TR>
								<%
										vTotalSaldo = vTotalSaldo + rsParc("PARC_ValorTotal") + vValorTaxaBoleto
										vTotalBoleto = vTotalBoleto + vValorTaxaBoleto
									else
								%>
								<TR>
									<TD align=middle><FONT style="FONT-SIZE: 8pt" face=Arial><% = rsParc("PARC_Numero")%> </FONT></TD>
									<TD align=middle><FONT style="FONT-SIZE: 8pt" face=Arial><% = rsParc("PARC_Vencimento")%> </FONT></TD>
									<TD align=right><FONT style="FONT-SIZE: 8pt" face=Arial><% if isnull(rsParc("PARC_TaxaBoleto")) then Response.Write FormatNumber(rsParc("PARC_ValorPagamento")) else Response.Write FormatNumber(rsParc("PARC_ValorPagamento") - rsParc("PARC_TaxaBoleto")) %> </FONT></TD>
									<TD align=right><FONT style="FONT-SIZE: 8pt" face=Arial><% if not isnull(rsParc("PARC_TaxaBoleto")) then Response.Write FormatNumber(rsParc("PARC_TaxaBoleto")) else Response.Write "0,00"%></FONT></TD>
									<TD align=right><FONT style="FONT-SIZE: 8pt" face=Arial><% = FormatNumber(rsParc("PARC_ValorPagamento")) %> </FONT></TD>
									<TD align=middle><FONT style="FONT-SIZE: 8pt" face=Arial>Baixado</FONT></TD>
									<TD align=right><FONT style="FONT-SIZE: 8pt" face=Arial>0,00 </FONT></TD>
								</TR>
								<%
										if not isnull(rsParc("PARC_TaxaBoleto")) then
											vTotalBoleto = vTotalBoleto + rsParc("PARC_TaxaBoleto")
										end if
									end if
									vTotalParcelas = vTotalParcelas + rsParc("PARC_ValorTotal")
									rsParc.MoveNext
								Loop
								%>
								<TR bgColor=#d3d3d3>
									<TD align=middle colSpan=2><FONT style="FONT-SIZE: 8pt" face=Arial><B>TOTAL</FONT></B></TD>
									<TD align=right><FONT style="FONT-SIZE: 8pt" face=Arial><B><% = FormatNumber(vTotalParcelas)%></FONT></B></TD>
									<TD align=right><FONT style="FONT-SIZE: 8pt" face=Arial><B><% = FormatNumber(vTotalBoleto)%></FONT></B></TD>
									<TD align=right><FONT style="FONT-SIZE: 8pt" face=Arial><B><% = FormatNumber(vTotalParcelas + vTotalBoleto)%></FONT></B></TD>
									<TD>&nbsp;</TD>
									<TD align=right><FONT style="FONT-SIZE: 8pt" face=Arial><B><% = FormatNumber(vTotalSaldo)%></FONT></B></TD>
								</TR>
							</TBODY>
							</TABLE>
						</TD>
					</TR>
				</TBODY>
				</TABLE>
			</TD>
		</TR>
	</TABLE>
	<TABLE borderColor=black cellSpacing=0 borderColorDark=white cellPadding=0 width=640 border=0 ID="Table8">
	<TBODY>
		<TR>
			<TD width="100%"><FONT style="FONT-SIZE: 8pt" face=Arial>Pague sua(s) parcela(s) em dia e em caso de dúvida entre em contato conosco através do telefone (11) 2165-9380 ou envie e-mail para <A title=mailto:atendimento@empresariocobranca.com.br href="mailto:atendimento@empresariocobranca.com.br">atendimento@empresariocobranca.com.br</A>, que teremos a maior satisfação em atendê-lo.</B> </FONT></td>
		</tr>
	</TBODY>
	</table> 
	<br>
	<TABLE WIDTH="640" CELLSPACING=0 CELLPADDING=0 BORDER=1 borderbolor=black ID="Table9">
  		<TR>
			<TD ALIGN=center VALIGN=BOTTOM><FONT class=ld><B>ESTE BOLETO QUITA A PARCELA Nº <% = rsParcelas("PARC_Numero")%> DO ACORDO</B></FONT></TD>
		</TR>
	</TABLE>
	<br>
	<TABLE WIDTH="640" CELLSPACING=0 CELLPADDING=0 BORDER=0 ID="Table32">
  		<TR>
			<TD class=cp VALIGN=BOTTOM WIDTH=225><IMG SRC="../images/logobanco_<% = rsBanco("BANC_ID_Boleto")%>.gif"></TD>
			<TD ALIGN=RIGHT VALIGN=BOTTOM><FONT class=ld><B>RECIBO DO SACADO</B></FONT></TD>
		</TR>
	</TABLE>
	<TABLE WIDTH="640" BORDER=1 CELLSPACING=0 CELLPADDING=1 ID="Table33">
	<TR>
				<TD COLSPAN=2><FONT class=ct>Cedente</FONT><BR><FONT class=cp>&nbsp;<%=cons_cedente%></FONT></TD>
				<TD width=15%><FONT class=ct>Ag&ecirc;ncia / C&oacute;digo Cedente</FONT><BR>
						<FONT align=center class=cn>&nbsp;<% = var_textocodigocedente %></FONT></TD>
   				<TD width=15%><FONT class=ct>Nosso Número</FONT><BR><FONT class=cn>&nbsp;<% = var_textonossonumero %></FONT></TD>
				<TD width=20% bgcolor="#CCCCCC">
				<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0 ID="Table34">
						<TR><TD align=left><FONT class=ct>Vencimento</FONT></TD></TR>
							<TR><TD align=center><FONT class=cp><%=var_datavencimento%></FONT></TD></TR>
					</TABLE>
				</TD>
		</TR>
		<TR>
   			<TD COLSPAN=2><FONT class=ct>Sacado</FONT><BR><FONT class=cp>&nbsp;<%=var_sacado%></FONT></TD>
			<TD width=15%><FONT class=ct>Parcela / Plano</FONT><BR><FONT align=center class=cn>&nbsp;<%= var_parcelaplano%></FONT></TD>
   			<TD width=15%><FONT class=ct>Número Documento</FONT><BR><FONT class=cn>&nbsp;<%=var_numerodoc%></FONT></TD>
				<TD width=20% bgcolor="#CCCCCC">
				<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0 ID="Table35">
						<TR><TD align=left><FONT class=ct>Valor do Documento </FONT></TD></TR>
							<TR><TD align=center><FONT class=cp><% if rsBanco("BANC_ID_Boleto") = 237 or vCONT_ID = 77 then Response.Write FormatNumber(var_valordocumento) else Response.Write FormatNumber(var_valordocumento - vValorTaxaBoleto)%></FONT></TD></TR>
					</TABLE>
				</TD>
		</TR>
		<TR>
   			<TD><FONT class=ct>Contrato</FONT><BR><FONT class=cp>&nbsp;<%if Session("CodigoCliente") = 23 and vCont_id = 53 then response.Write FormataContrato(rsParcelas("CTRA_Numero"))  & " - " & vDescCert   else response.Write var_contrato  end if%></FONT></TD>
   			<TD width=15%><FONT class=ct>(-) Desconto / Abatimento</FONT><BR><FONT class=cp><br></FONT></TD>
   			<TD width=15%><FONT class=ct>(+) Mora / Multa</FONT><BR><FONT class=cn><br></FONT></TD>
   			<TD width=15%><FONT class=ct>(+) Outros Acréscimos</FONT><BR><FONT class=cn>&nbsp;<% if rsBanco("BANC_ID_Boleto") = 237 or vCONT_ID = 77  then Response.Write "&nbsp;" else Response.Write FormatNumber(vValorTaxaBoleto)%></FONT></TD>
   			<TD width=20% bgcolor="#CCCCCC"><FONT class=ct>(=) Valor Cobrado</FONT><BR><FONT class=cp><center><%if vJurosAtraso or rsBanco("BANC_ID_Boleto") = 237 then response.Write "&nbsp;" else Response.Write FormatNumber(var_valordocumento)%></center></FONT></TD>
		</TR>
	</TABLE>
	<TABLE WIDTH="640" CELLSPACING=0 CELLPADDING=0 BORDER=0 ID="Table36">
		<TR>
				<TD align=right><FONT class=ct>Autentica&ccedil;&atilde;o Mec&acirc;nica</FONT><BR></TD>
	</TR>
	</TABLE>
	<br>

<%
end sub

'Esta função mostra as informações de decrição de boleto para o escritorio empresario para o contratante
'ZOGBI, cujo cont_id pode ser 100 ou 104
sub DescBoletoEmpresarioZogbi()
%>
	<TABLE WIDTH="640" CELLSPACING=0 CELLPADDING=0 BORDER=0 ID="Table12">
  		<TR>
			<TD VALIGN=top WIDTH=225><IMG SRC="../images/logo_cliente_logon_25.gif"></TD>
			<TD ALIGN=RIGHT VALIGN=bottom class="cn5">São Paulo, <%=Day(date)%> de <%=MonthName(Month(date))%> de <%=Year(date)%>	</TD>
		</TR>
		<tr>
			<td colspan=2 ALIGN=left><br>	
				<TABLE WIDTH="460" ALIGN=left CELLSPACING=0 CELLPADDING=0 BORDER=0 ID="Table13">
					<tr>
						<td class=cn5> Prezado(a) Sr.(a) </td>
					</tr>
					<tr>
						<td class=cn5><b><%=var_sacado%>  </b></td>
					</tr>
					<tr>
						<td class=cn5><b> Contrato: <%=FormataContrato(rsParcelas("CTRA_Numero"))%>  </b></td>
					</tr>
					<tr>
						<td class=cn5>&nbsp;</td>
					</tr>
					<tr>
						<td class=cn5><b>Credor: Lojas Zogbi</b></td>
					</tr>
				</TABLE>
			</td>
		</tr>
		
		<tr>
			<td colspan=2 ALIGN=left class=cn5>	
			<br>
					Conforme acordo firmado anteriormente, segue o demonstrativo dos débitos e o boleto bancário para a quitação do débito em aberto, conforme o acordo negociado.
					O não pagamento até a data do vencimento acarretará no cancelamento do acordo e nova inclusão de seu nome no órgão de proteção ao crédito (SPC).						
			</td>
		</tr>
		
		<tr>
			<td colspan=2 align=center>
					<br>
					<TABLE borderColor=black cellSpacing=0 borderColorDark=white cellPadding=0 width="90%" border=1 ID="Table3">
					<TBODY>
						<TR bgColor=#d3d3d3>
							<TD align=middle width="100%" colSpan=5><FONT style="FONT-SIZE: 10pt" face=Arial><B>DEMONSTRATIVO DOS VALORES DO CONTRATO ORIGINAL EM ABERTO</B></FONT></TD>
						</TR>
						<TR style="FONT-WEIGHT: normal; FONT-SIZE: 8pt; LINE-HEIGHT: normal; FONT-STYLE: normal; FONT-VARIANT: normal" align=middle bgColor=#d3d3d3><FONT face=Arial>
							<TD>Parcela</TD>
							<TD>Vencimento</TD>
							<TD>Original</TD>
							<TD>Negociado</TD>
							<TD>Total</TD>
						</TR>
						<%
						
						vValorPrincipalAtual = 0
						vValorCorrigido = 0
						Set rsTit = Bd.Execute("SELECT * FROM Transacoes_de_Acordo t WITH (NOLOCK) LEFT JOIN Tipos_de_Documento td WITH (NOLOCK) ON t.TDOC_ID = td.TDOC_ID WHERE ACOR_ID = " & rsParcelas("ACOR_ID"))
						if rsTit.EOF then
							Set rsTit = Bd.Execute("SELECT * FROM Transacoes t WITH (NOLOCK) LEFT JOIN Tipos_de_Documento td WITH (NOLOCK) ON t.TDOC_ID = td.TDOC_ID WHERE CTRA_ID = " & rsParcelas("CTRA_ID"))
						end if
						Do While not rsTit.EOF
							if rsTit("TDOC_ID") = "" or isnull(rsTit("TDOC_ID")) then
								vValorPrincipalAtual = vValorPrincipalAtual + rsTit("TRAN_Valor")
								if rsTit("TRAN_Vencimento") > rsParcelas("ACOR_Data") then
									vValorCorrigido = vValorCorrigido
								else
									vValorCorrigido = vValorCorrigido + AtualizaValor(rsParcelas("CART_ID"), rsTit("TRAN_Vencimento"), rsParcelas("ACOR_Data"), rsTit("TRAN_Valor"), "+", 1, rsParcelas("Plano"), rsParcelas("CTRA_DataRecebimentoContrato"), -1, rsTit("TRAN_DataRecebimento"), rsTit("TRAN_Vencimento"))
								end if
							else
								if rsTit("TRAN_Vencimento") > rsParcelas("ACOR_Data") then
									vValorCorrigido = vValorCorrigido
								else
									vValorCorrigido = eval(CStr(Replace(Replace(vValorCorrigido,".",""),",",".")) & CStr(rsTit("TDOC_Sinal")) & CStr(Replace(Replace(AtualizaValor(rsParcelas("CART_ID"), rsTit("TRAN_Vencimento"), rsParcelas("ACOR_Data"), rsTit("TRAN_Valor"), rsTit("TDOC_Sinal"), 1, rsParcelas("Plano"), rsParcelas("CTRA_DataRecebimentoContrato"), -1, rsTit("TRAN_DataRecebimento"), rsTit("TRAN_Vencimento")),".",""),",",".")))
								end if
								vValorPrincipalAtual = eval(CStr(Replace(Replace(vValorPrincipalAtual,".",""),",",".")) & CStr(rsTit("TDOC_Sinal")) & CStr(Replace(Replace(rsTit("TRAN_Valor"),".",""),",",".")))
							end if
							rsTit.MoveNext
						loop
														
						vPercentDesc = 1 - (rsParcelas("ACOR_ValorAcordo") / vValorCorrigido)
						Set rsTit = Bd.Execute("SELECT * FROM Transacoes_de_Acordo t WITH (NOLOCK) LEFT JOIN Tipos_de_Documento td WITH (NOLOCK) ON t.TDOC_ID = td.TDOC_ID WHERE ACOR_ID = " & rsParcelas("ACOR_ID"))
						if rsTit.EOF then
							Set rsTit = Bd.Execute("SELECT * FROM Transacoes t WITH (NOLOCK) LEFT JOIN Tipos_de_Documento td WITH (NOLOCK) ON t.TDOC_ID = td.TDOC_ID WHERE CTRA_ID = " & rsParcelas("CTRA_ID"))
						end if
						vTotalOriginal = 0
						vTotalNegociado = 0
						Do while not rsTit.EOF
							if isnull(rsTit("TDOC_Sinal")) then
								vValAt = AtualizaValor(rsParcelas("CART_ID"), rsTit("TRAN_Vencimento"), rsParcelas("ACOR_Data"), rsTit("TRAN_Valor"), "+", 1, rsParcelas("Plano"), rsParcelas("CTRA_DataRecebimentoContrato"), -1, rsTit("TRAN_DataRecebimento"), rsTit("TRAN_Vencimento"))
							else
								vValAt = AtualizaValor(rsParcelas("CART_ID"), rsTit("TRAN_Vencimento"), rsParcelas("ACOR_Data"), rsTit("TRAN_Valor"), rsTit("TDOC_Sinal"), 1, rsParcelas("Plano"), rsParcelas("CTRA_DataRecebimentoContrato"), -1, rsTit("TRAN_DataRecebimento"), rsTit("TRAN_Vencimento"))
							end if
							vValorOriginal = vValAt
							vValDesc = Round(vValAt * vPercentDesc, 2)
							vValorLiquido = Round(vValAt - vValDesc, 2)
							%>
							<TR>
								<TD align=middle><FONT style="FONT-SIZE: 8pt" face=Arial><% = rsTit("TRAN_NumTitulo")%></FONT></TD>
								<TD align=middle><FONT style="FONT-SIZE: 8pt" face=Arial><% = rsTit("TRAN_Vencimento")%></FONT></TD>
								<TD align=right><FONT style="FONT-SIZE: 8pt" face=Arial><% = rsTit("TDOC_Sinal") & FormatNumber(vValAt)%></FONT></TD>
								<TD align=right><FONT style="FONT-SIZE: 8pt" face=Arial><% = rsTit("TDOC_Sinal") & FormatNumber(vValorLiquido)%></FONT></TD>
								<TD align=right><FONT style="FONT-SIZE: 8pt" face=Arial><% = rsTit("TDOC_Sinal") & FormatNumber(vValorLiquido)%></FONT></TD>
							</TR>
							<%
							if rsTit("TDOC_Sinal") = "-" then
								vTotalOriginal = vTotalOriginal - vValorOriginal
								vTotalNegociado = vTotalNegociado - vValorLiquido
							else
								vTotalOriginal = vTotalOriginal + vValorOriginal
								vTotalNegociado = vTotalNegociado + vValorLiquido
							end if
							rsTit.MoveNext
						Loop

						%>
						<TR bgColor=#d3d3d3>
							<TD align=middle colSpan=2><FONT style="FONT-SIZE: 8pt" face=Arial><B>TOTAL</FONT></B></TD>
							<TD align=right><FONT style="FONT-SIZE: 8pt" face=Arial><B><% = vTotalOriginal%></FONT></B></TD>
							<TD align=right><FONT style="FONT-SIZE: 8pt" face=Arial><B><% = vTotalNegociado%></FONT></B></TD>
							<TD align=right><FONT style="FONT-SIZE: 8pt" face=Arial><B><% = vTotalNegociado%></FONT></B></TD>
						</TR>
					</TBODY>
					</TABLE>
					<br>
					
					<TABLE borderColor=black cellSpacing=0 borderColorDark=white cellPadding=0 width="90%" border=1 ID="Table10">
					<TBODY>
						<TR bgColor=#d3d3d3>
							<TD align=middle width="100%" colSpan=7><FONT style="FONT-SIZE: 10pt" face=Arial><B>DEMONSTRATIVO DAS PARCELAS DO ACORDO FIRMADO EM <% = rsParcelas("ACOR_Data")%></B></FONT></TD>
						</TR>
						<TR bgColor=#d3d3d3>
							<TD align=middle><FONT style="FONT-SIZE: 8pt" face=Arial> Nº </FONT></TD>
							<TD align=middle><FONT style="FONT-SIZE: 8pt" face=Arial> Vencimento </FONT></TD>
							<TD align=middle><FONT style="FONT-SIZE: 8pt" face=Arial> Parcela </FONT></TD>
							<TD align=middle><FONT style="FONT-SIZE: 8pt" face=Arial>IOF + Boleto</FONT></TD>
							<TD align=middle><FONT style="FONT-SIZE: 8pt" face=Arial> Total </FONT></TD>
							<TD align=middle><FONT style="FONT-SIZE: 8pt" face=Arial> Status </FONT></TD>
							<TD align=middle><FONT style="FONT-SIZE: 8pt" face=Arial> Saldo </FONT></TD>
						</TR>
						<%
						vTotalParcelas = 0
						vTotalBoleto = 0
						vTotalSaldo = 0
						Set rsParc = BD.Execute("SELECT * FROM Parcelas WITH (NOLOCK) WHERE ACOR_ID = " & rsParcelas("ACOR_ID") & " ORDER BY PARC_Numero")
						Do While not rsParc.EOF
							if IsNull(rsParc("PARC_DataPagamento")) then 
						%>
						<TR>
							<TD align=middle><FONT style="FONT-SIZE: 8pt" face=Arial><% = rsParc("PARC_Numero")%> </FONT></TD>
							<TD align=middle><FONT style="FONT-SIZE: 8pt" face=Arial><% = rsParc("PARC_Vencimento")%> </FONT></TD>
							<TD align=right><FONT style="FONT-SIZE: 8pt" face=Arial><% = FormatNumber(rsParc("PARC_ValorTotal")) %> </FONT></TD>
							<TD align=right><FONT style="FONT-SIZE: 8pt" face=Arial><% = FormatNumber(vValorTaxaBoleto)%></FONT></TD>
							<TD align=right><FONT style="FONT-SIZE: 8pt" face=Arial><% = FormatNumber(rsParc("PARC_ValorTotal") + vValorTaxaBoleto) %> </FONT></TD>
							<TD align=middle><FONT style="FONT-SIZE: 8pt" face=Arial>Aberto</FONT></TD>
							<TD align=right><FONT style="FONT-SIZE: 8pt" face=Arial><% = FormatNumber(rsParc("PARC_ValorTotal") + vValorTaxaBoleto) %> </FONT></TD>
						</TR>
						<%
								vTotalSaldo = vTotalSaldo + rsParc("PARC_ValorTotal") + vValorTaxaBoleto
								vTotalBoleto = vTotalBoleto + vValorTaxaBoleto
							else
						%>
						<TR>
							<TD align=middle><FONT style="FONT-SIZE: 8pt" face=Arial><% = rsParc("PARC_Numero")%> </FONT></TD>
							<TD align=middle><FONT style="FONT-SIZE: 8pt" face=Arial><% = rsParc("PARC_Vencimento")%> </FONT></TD>
							<TD align=right><FONT style="FONT-SIZE: 8pt" face=Arial><% if isnull(rsParc("PARC_TaxaBoleto")) then Response.Write FormatNumber(rsParc("PARC_ValorPagamento")) else Response.Write FormatNumber(rsParc("PARC_ValorPagamento") - rsParc("PARC_TaxaBoleto")) %> </FONT></TD>
							<TD align=right><FONT style="FONT-SIZE: 8pt" face=Arial><% if not isnull(rsParc("PARC_TaxaBoleto")) then Response.Write FormatNumber(rsParc("PARC_TaxaBoleto")) else Response.Write "0,00"%></FONT></TD>
							<TD align=right><FONT style="FONT-SIZE: 8pt" face=Arial><% = FormatNumber(rsParc("PARC_ValorPagamento")) %> </FONT></TD>
							<TD align=middle><FONT style="FONT-SIZE: 8pt" face=Arial>Baixado</FONT></TD>
							<TD align=right><FONT style="FONT-SIZE: 8pt" face=Arial>0,00 </FONT></TD>
						</TR>
						<%
								if not isnull(rsParc("PARC_TaxaBoleto")) then
									vTotalBoleto = vTotalBoleto + rsParc("PARC_TaxaBoleto")
								end if
							end if
							vTotalParcelas = vTotalParcelas + rsParc("PARC_ValorTotal")
							rsParc.MoveNext
						Loop
						%>
						<TR bgColor=#d3d3d3>
							<TD align=middle colSpan=2><FONT style="FONT-SIZE: 8pt" face=Arial><B>TOTAL</FONT></B></TD>
							<TD align=right><FONT style="FONT-SIZE: 8pt" face=Arial><B><% = FormatNumber(vTotalParcelas)%></FONT></B></TD>
							<TD align=right><FONT style="FONT-SIZE: 8pt" face=Arial><B><% = FormatNumber(vTotalBoleto)%></FONT></B></TD>
							<TD align=right><FONT style="FONT-SIZE: 8pt" face=Arial><B><% = FormatNumber(vTotalParcelas + vTotalBoleto)%></FONT></B></TD>
							<TD>&nbsp;</TD>
							<TD align=right><FONT style="FONT-SIZE: 8pt" face=Arial><B><% = FormatNumber(vTotalSaldo)%></FONT></B></TD>
						</TR>
					</TBODY>
					</TABLE>
					
			</td>
		</tr>
		<tr>
			<td colspan=2 class="cn5">
			<br>
					Para informações, contate-nos: Ligue <b> (11) 2165 -9436 </b>e fale com nossa central de atendimento de Segunda a Sexta das 08h00 às 20h30 e aos sábados das 08h00 às 18h00, ou envie e-mail para atendimento@empresariocobranca.com.br .			
			</td>
		</tr>
		<tr>
			<td colspan=2 class="cn5" align=center>
				<br>
				<b>	Este documento é válido para os pagamentos efetuados até o dia <%=var_datavencimento%>. </b>
			</td>
		</tr>
		</table>
<%
end sub
%>

<%
'http://localhost/teste/boleto_laser.asp?avulso=sim&cart_id=1001&cont_id=1&recuperador=42001&ctra_id=37095&tipoboleto=2
'http://localhost/teste/boleto_laser.asp?avulso=sim&cart_id=2001&cont_id=1&recuperador=18001&ctra_id=53095&tipoboleto=2
'http://localhost/teste/boleto_laser.asp?avulso=sim&cart_id=1001&cont_id=1&recuperador=17001&ctra_id=26099&tipoboleto=2

Response.AddHeader "pragma","no-cache"
Response.AddHeader "cache-control","no-store"
Response.CacheControl = "no-cache"
Response.Expires = -100000

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
<!-- #include file="func\boleto_funcoes.asp" -->
<!-- #include file="func\funcoes.asp" -->
<!-- #include file="func\log_operacoes.asp" -->
<!-- #include file="desc_boleto_laser.asp" -->
<body>

<div id="aguarde" style="position:absolute; width:100%; left: 0px; top: 0px; overflow: auto;">
<br>
<br>
<br>
<br>
	<table border=0 cellspacing=0 cellpadding=0 align=center ID="Table17">
	  <tr>
		 
      <td valign=top class=texto1>&nbsp;</td>
	  </tr>
	</table>
</div>
<CENTER>
	<TABLE WIDTH="640" CELLSPACING=0 CELLPADDING=0 BORDER=0 ID="Table19">
  		<TR>
			<TD class=cp VALIGN=BOTTOM WIDTH=225><IMG SRC="images/logobanco_.gif"></TD>
			<TD ALIGN=RIGHT VALIGN=BOTTOM><FONT class=ld><B>RECIBO DO SACADO</B></FONT></TD>
		</TR>
	</TABLE>
	<TABLE WIDTH="640" BORDER=1 CELLSPACING=0 CELLPADDING=1 ID="Table20">
		<TR>
   			<TD COLSPAN=2><FONT class=ct>Sacado</FONT><BR>
        <FONT class=cp>&nbsp;</FONT></TD>
   			<TD width=15%><FONT class=ct>Número Documento</FONT><BR>
        <FONT class=cn>&nbsp;</FONT></TD>
			<TD width=20% bgcolor="#CCCCCC">
			<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0 ID="Table21">
					<TR><TD align=left><FONT class=ct>Vencimento</FONT></TD></TR>
						<TR>
            <TD align=center>&nbsp;</TD>
          </TR>
				</TABLE>
			</TD>
			</TD>
		</TR>
	</TABLE>
	<table width="640" cellpadding="3" cellspacing="0" align="center" border=0 bordercolor=black height=460 ID="Table1">
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
					<td class=texto1>O pagamento do boleto não implica a reativação de conta corrente e limites de crédito, cuja análise estará sujeita ao .</td>
				</tr>
				<TR>
					<td class=texto1>Caso a situação já tenha sido regularizada, favor desconsiderar este documento.</td>
				</tr>
			</table>
			<br>
			<table width="620" cellpadding="0" cellspacing="0" align="center" border=0 ID="Table25">
				<TR>
					<td class=texto1>Valor total do acordo: </td>
				</tr>
				<TR>
					<td class=texto1>Quantidade de parcelas do acordo:  parcela(s)</td>
				</tr>
				<TR>
					
					<td class=texto1>Número da parcela do acordo: </td>
				</tr>
				<TR>
					<td class=texto1>Valor da parcela: </td>
				</tr>
				<TR>
					<td class=texto1>Código negociador: </td>
				</tr>
			</table> 
			<br>
			<table width="620" cellpadding="0" cellspacing="0" align="center" border=0 ID="Table27">
				<TR>
					<td class=texto1>Débito referente ao(s) contrato(s)/parcela(s):</td>
				</tr>
				<TR>
					<td class=texto1></td>
				</tr>
				<tr>
					<td class=texto1>
		Informamos que até o momento não acusamos o pagamento de sua parcela vencida em  .
<br>
		Solicitamos a regularização com urgência para evitarmos o cancelamento do acordo e a inclusão nos orgãos de proteção ao crédito.
<br>
		NÃO PERCA AS FACILIDADES E DESCONTOS CONCEDIDOS , EFETUE O PAGAMENTO DE SUA PARCELA EM ATRASO O MAIS BREVE POSSÍVEL. 
<br>
	    * Caso o pagamento tenha sido efetuado , por favor desconsiderar aviso .
			</td>
				</tr>
				
			</table> 
		</td>
	</tr> 
	</table> 
	</TABLE>
    <table width="640" cellpadding="3" cellspacing="0" align="center" border=0 bordercolor=black ID="Table3">
     <tr>  
      	
      <td class=texto1><b>Em caso de dúvida ligue para ou compareça no seguinte 
        endereço:</b><br>
        <br>
        <b>De 2ª a 6ª feira: de 9:00 hs às 18:00 hs</b></td>
     </tr>
    </table>
    <br>
	<TABLE WIDTH="640" BORDER=1 CELLSPACING=0 CELLPADDING=1 ID="Table24">
		<TR>
   			<TD><FONT class=ct>Código do Documento</FONT><BR>
        <FONT class=cn>&nbsp;</FONT></TD>
			<TD valign=top><FONT class=ct>Espécie</FONT><BR>
        <FONT class=cn>&nbsp;</FONT></TD>
			<TD valign=top><FONT class=ct>Quantidade</FONT><BR><FONT class=cn>&nbsp;</FONT></TD>
			<TD bgcolor="#CCCCCC">
				<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0 ID="Table22">
					<TR><TD align=left><FONT class=ct>Valor do Documento </FONT></TD></TR>
						<TR>
            <TD align=center>&nbsp;</TD>
          </TR>
				</TABLE>
			</TD>
			<TD valign=top><FONT class=ct>Espécie Doc.</FONT><BR><FONT class=cn>&nbsp;</FONT></TD>
			<TD><FONT class=ct>C&oacute;digo Cedente</FONT><BR>
        <FONT align=center class=cn>&nbsp;</FONT></TD>
		</TR>
	</TABLE>
	<TABLE WIDTH="640" CELLSPACING=0 CELLPADDING=0 BORDER=0 ID="Table23">
		<TR>
				<TD align=right><FONT class=ct>Autentica&ccedil;&atilde;o Mec&acirc;nica</FONT><BR></TD>
	</TR>
	</TABLE>
	<br>
  <img src="images/linha2.gif" border=0 width="640" height=1> 
  <TABLE WIDTH="640" CELLSPACING=0 CELLPADDING=0 BORDER=0 ID="Table10">
  		<TR>
			<TD class=cp VALIGN=BOTTOM WIDTH=225><IMG SRC="images/logobanco_.gif"></TD>
			<TD ALIGN=RIGHT VALIGN=BOTTOM><FONT class=ld><B>RECIBO DO SACADO</B></FONT></TD>
		</TR>
	</TABLE>
	<TABLE WIDTH="640" BORDER=1 CELLSPACING=0 CELLPADDING=1 ID="Table11">
	<TR>
				<TD COLSPAN=2><FONT class=ct>Cedente</FONT><BR>
        <FONT class=cp>&nbsp;</FONT></TD>
				<TD width=15%><FONT class=ct>Ag&ecirc;ncia / C&oacute;digo Cedente</FONT><BR>
        <FONT align=center class=cn>&nbsp;</FONT></TD>
   				<TD width=15%><FONT class=ct>Nosso Número</FONT><BR>
        <FONT class=cn>&nbsp;</FONT></TD>
				<TD width=20% bgcolor="#CCCCCC">
				<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0 ID="Table12">
						<TR><TD align=left><FONT class=ct>Vencimento</FONT></TD></TR>
							<TR>
            <TD align=center>&nbsp;</TD>
          </TR>
					</TABLE>
				</TD>
		</TR>
		<TR>
   			<TD width=15%><FONT class=ct>Número Documento</FONT><BR>
        <FONT class=cn>&nbsp;</FONT></TD>
				<TD width=20% bgcolor="#CCCCCC">
				<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0 ID="Table13">
						<TR><TD align=left><FONT class=ct>Valor do Documento </FONT></TD></TR>
							<TR>
            <TD align=center>&nbsp;</TD>
          </TR>
					</TABLE>
				</TD>
		</TR>
		<TR>
   			<TD><FONT class=ct>Contrato</FONT><BR>
        <FONT class=cp>&nbsp;</FONT></TD>
   			<TD width=15%><FONT class=ct>(-) Desconto / Abatimento</FONT><BR><FONT class=cp><br></FONT></TD>
   			<TD width=15%><FONT class=ct>(+) Mora / Multa</FONT><BR><FONT class=cn><br></FONT></TD>
   			<TD width=15%><FONT class=ct>(+) Outros Acréscimos</FONT><BR>
        <FONT class=cn>&nbsp;</FONT></TD>
   			<TD width=20% bgcolor="#CCCCCC"><FONT class=ct>(=) Valor Cobrado</FONT><BR><FONT class=cp><center>
        </center></FONT></TD>
		</TR>
	</TABLE>
	<TABLE WIDTH="640" CELLSPACING=0 CELLPADDING=0 BORDER=0 ID="Table14">
		<TR>
				<TD align=right><FONT class=ct>Autentica&ccedil;&atilde;o Mec&acirc;nica</FONT><BR></TD>
	</TR>
	</TABLE>
	
  <img src="images/corte.gif" border=0 width="640"><br>
	<TABLE WIDTH="640" BORDER=0 CELLSPACING=0 CELLPADDING=0 ID="Table15">
	<tr>
			<td class=cp width=150><div align="left"><img src="images/logobanco_.gif"></div></td>
  			<td width=3 valign="bottom"><img height=22 src="images/barra.gif" width=2 border=0></td>
	  		<td class=cpt  width=58 valign="bottom"><div align="center"><font class="bc">-</font></div></td>
  			<td width=3 valign="bottom"><img height=22 src="images/barra.gif" width=2 border=0></td>
	  		<td class=ld align=right width=453 valign="bottom"><span class='ld'>
        <p align="right">&nbsp;</span></td>
	</tr>
	</TABLE>
	<TABLE WIDTH="640" BORDER=1 CELLSPACING=0 CELLPADDING=1 ID="Table16">
	<TR>
				<TD COLSPAN=5 WIDTH=500>
						<FONT class=ct>Local de Pagamento</FONT><BR>
        <FONT class=cp>&nbsp;</FONT> </TD>
				<TD bgcolor="#CCCCCC">
				<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0 ID="Table28">
						<TR><TD align=left><FONT class=ct>Vencimento</FONT></TD></TR>
						<TR>
            <TD align=center>&nbsp;</TD>
          </TR>
					</TABLE>
				</TD>
		</TR>
		<TR>
				<TD COLSPAN=5 WIDTH=500><FONT class=ct>Cedente</FONT><BR>
        <FONT class=cn>&nbsp;</FONT></TD>
				<TD width=170>
				<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0 ID="Table29">
						<TR><TD align=left><FONT class=ct>Ag&ecirc;ncia / C&oacute;digo Cedente</FONT></TD></TR>
							<TR>
            <TD align=center>&nbsp;</TD>
          </TR>
					</TABLE>
				</TD>
		</TR>
		<TR>
				<TD valign=top><FONT class=ct>Data Documento</FONT><BR>
        <FONT class=cn>&nbsp;</FONT></TD>
				<TD valign=top><FONT class=ct>Número Documento</FONT><BR>
        <FONT class=cn>&nbsp;</FONT></TD>
				<TD valign=top><FONT class=ct>Tipo Docu.</FONT><BR>
        <FONT class=cn>&nbsp;</FONT></TD>
				<TD valign=top><FONT class=ct>Aceite</FONT><BR>
        <FONT class=cn>&nbsp;</FONT></TD>
				<TD valign=top><FONT class=ct>Data Processamento</FONT><BR>
        <FONT class=cn>&nbsp;</FONT></TD>
				<TD width=170>
				<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0 ID="Table30">
						<TR><TD align=left><FONT class=ct>Nosso Número</FONT></TD></TR>
							<TR>
            <TD align=center>&nbsp;</TD>
          </TR>
</TABLE>
				</TD>
		</TR>
		<TR>
				<TD valign=top><FONT class=ct>Uso Banco</FONT><BR><FONT class=cn>&nbsp;</FONT></TD>
				<TD valign=top><FONT class=ct>Carteira</FONT><BR>
        <FONT class=cn>&nbsp;</FONT></TD>
				<TD valign=top><FONT class=ct>Espécie</FONT><BR>
        <FONT class=cn>&nbsp;</FONT></TD>
				<TD valign=top><FONT class=ct>Quantidade</FONT><BR><FONT class=cn>&nbsp;</FONT></TD>
				<TD valign=top><FONT class=ct>Valor</FONT><BR><FONT class=cn>&nbsp;</FONT></TD>
				<TD width=170 bgcolor="#CCCCCC">
				<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0 ID="Table37">
						<TR><TD align=left><FONT class=ct>Valor do Documento </FONT></TD></TR>
							<TR>
            <TD align=center><FONT class=cp>&nbsp; </FONT></TD>
          </TR>
					</TABLE>
				</TD>
		</TR>
		<TR>
				<TH COLSPAN=5 ROWSPAN=4 valign=top align=LEFT ><FONT class=ct>Instru&ccedil;&otilde;es (Todas as informações deste bloqueto são de inteira responsabilidade do cedente)</FONT><BR>
 					<TABLE WIDTH="475" ALIGN=RIGHT CELLSPACING=0 CELLPADDING=0 BORDER=0 ID="Table38">
						<TR>
							
            <TD valign=top align=left> <FONT class=cn> <br>
              <b>Sr. Cliente Caso tenha ocorrido mudança de endereço, favor comparecer 
              à agência para atualização dos dados cadastrais para que o carnê 
              com os próximos pagamentos chegue no endereço correto.</b> </FONT> 
            </TD>
						</TR>
					</TABLE>
				</TH>
				<TD WIDTH=170><FONT class=ct>(-) Desconto / Abatimento</FONT><BR><FONT class=cn3>&nbsp;</FONT></TD>
		</TR>
		<TR>
				<TD WIDTH=170><FONT class=ct>(+) Mora / Multa</FONT><BR><FONT class=cn3>&nbsp;</FONT></TD>
		</TR>
		<TR>
				<TD WIDTH=170 ><FONT class=ct>(+) Outros Acréscimos</FONT><BR><FONT class=cn><center>
          &nbsp; 
        </center></FONT></TD>
		</TR>
		<TR>
				<TD WIDTH=170  bgcolor="#CCCCCC">
					<TABLE WIDTH=100% BORDER=0 CELLSPACING=0 CELLPADDING=0 ID="Table39">
							<TR><TD align=left><FONT class=ct>(=) Valor Cobrado</FONT></TD></TR>
								<TR>
            <TD align=center>&nbsp;</TD>
          </TR>
					</TABLE>
				</TD>
		</TR>
		<TR>
				<TD valign=top>
							<TABLE WIDTH="638" ALIGN=RIGHT CELLSPACING=0 CELLPADDING=0 BORDER=0 ID="Table40">
								<TR>
									<TD valign=top align=left width=100><FONT class=ct>Sacado</FONT></td>
									
            <TD valign=top align=left> <FONT class=cn5>&nbsp; </FONT> </TD>
								</TR>
							</TABLE>
				<TD  valign=top>
							<TABLE WIDTH="638" ALIGN=RIGHT CELLSPACING=0 CELLPADDING=0 BORDER=0 ID="Table44">
								<TR>
									<TD valign=top align=left width=100><FONT class=ct>Sacado </FONT></td>
									
            <TD valign=top align=left> <FONT class=cn5>&nbsp; </FONT> </TD>
								</TR>
							</TABLE>
				
				<TD valign=top>
						<FONT class=ct>Sacado</FONT><BR>
							<TABLE WIDTH="560" ALIGN=RIGHT CELLSPACING=0 CELLPADDING=0 BORDER=0 ID="Table41">
								<TR>
									
            <TD valign=top align=left> <FONT class=cn>&nbsp; </FONT> </TD>
								</TR>
							</TABLE>
				
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
							call wbarcode("23791272000000102404130060007222400500201090")
							'response.Write GeraBarraTexto("23791272000000102404130060007222400500201090")
						%>
				</TD>
		</TR>
	</TABLE>
</CENTER>
</body>
</HTML>
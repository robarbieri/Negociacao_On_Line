<%

'**************************
FUNCTION linhadigitavel_caixa(codigobarras)
'**************************
	cmplivre = mid(codigobarras,20,25)
	campo1 = left(codigobarras,4) & mid(cmplivre,1,5)
	campo1 = campo1 & calcdig10(campo1)
	campo1 = mid(campo1,1,5) & "." & mid(campo1,6,5)

	campo2 = mid(cmplivre,6,10)
	campo2 = campo2 & calcdig10(campo2)
	campo2 = mid(campo2,1,5) & "." & mid(campo2,6,6)

	campo3 = mid(cmplivre,16,10)
	campo3 = campo3 & calcdig10(campo3)
	campo3 = mid(campo3,1,5) & "." & mid(campo3,6,6)

	campo4 = mid(codigobarras,5,1)

	campo5 = mid(codigobarras,6,14)

	if campo5 = 0 then
		campo5 = "000"
	end if

	linhadigitavel_caixa = campo1 & "&nbsp;&nbsp;" & campo2 & "&nbsp;&nbsp;" & campo3 & "&nbsp;&nbsp;" & campo4 & "&nbsp;&nbsp;" & campo5
'*************************
END FUNCTION
'*************************

'**************************
FUNCTION linhadigitavel_banespa(codigobarras)
'**************************
	cmplivre = mid(codigobarras,20,25)
	campo1 = left(codigobarras,4) & mid(cmplivre,1,5)
	campo1 = campo1 & calcdig10(campo1)
	campo1 = mid(campo1,1,5) & "." & mid(campo1,6,5)

	campo2 = mid(cmplivre,6,10)
	campo2 = campo2 & calcdig10(campo2)
	campo2 = mid(campo2,1,5) & "." & mid(campo2,6,6)

	campo3 = mid(cmplivre,16,10)
	campo3 = campo3 & calcdig10(campo3)
	campo3 = mid(campo3,1,5) & "." & mid(campo3,6,6)

	campo4 = mid(codigobarras,5,1)

	campo5 = mid(codigobarras,6,14)

	if campo5 = 0 then
		campo5 = "000"
	end if

	linhadigitavel_banespa = campo1 & "&nbsp;&nbsp;" & campo2 & "&nbsp;&nbsp;" & campo3 & "&nbsp;&nbsp;" & campo4 & "&nbsp;&nbsp;" & campo5
'*************************
END FUNCTION
'*************************

'**************************
FUNCTION linhadigitavel_santander(codigobarras)
'**************************
	campo1 = left(codigobarras,4) & mid(codigobarras,20,5)
	campo1 = campo1 & calcdig10(campo1)
	campo1 = mid(campo1,1,5) & "." & mid(campo1,6,5)

	campo2 = mid(codigobarras,25,3) & mid(codigobarras,28,7)
	campo2 = campo2 & calcdig10(campo2)
	campo2 = mid(campo2,1,5) & "." & mid(campo2,6,6)

	campo3 = mid(codigobarras,35,6) & mid(codigobarras,41,4)
	campo3 = campo3 & calcdig10(campo3)
	campo3 = mid(campo3,1,5) & "." & mid(campo3,6,6)

	campo4 = mid(codigobarras,5,1)

	campo5 = int(mid(codigobarras,6,14))

	if campo5 = 0 then
		campo5 = "000"
	end if

	linhadigitavel_santander = campo1 & "&nbsp;&nbsp;" & campo2 & "&nbsp;&nbsp;" & campo3 & "&nbsp;&nbsp;" & campo4 & "&nbsp;&nbsp;" & campo5
'*************************
END FUNCTION
'*************************

'**************************
FUNCTION linhadigitavel_itau(codigobarras)
'**************************
	cmplivre = mid(codigobarras,20,25)
	campo1 = left(codigobarras,4) & mid(cmplivre,1,5)
	campo1 = campo1 & calcdig10(campo1)
	campo1 = mid(campo1,1,5) & "." & mid(campo1,6,5)

	campo2 = mid(cmplivre,6,10)
	campo2 = campo2&calcdig10(campo2)
	campo2 = mid(campo2,1,5) & "." & mid(campo2,6,6)

	campo3 = mid(cmplivre,16,10)
	campo3 = campo3 & calcdig10(campo3)
	campo3 = mid(campo3,1,5) & "." & mid(campo3,6,6)

	campo4 = mid(codigobarras,5,1)

	campo5 = int(mid(codigobarras,6,14))

	if campo5 = 0 then
		campo5 = "000"
	end if

	linhadigitavel_itau = campo1 & "&nbsp;&nbsp;" & campo2 & "&nbsp;&nbsp;" & campo3 & "&nbsp;&nbsp;" & campo4 & "&nbsp;&nbsp;" & campo5
'*************************
END FUNCTION
'*************************

'**************************
FUNCTION linhadigitavel_bradesco(codigobarras)
'**************************
	cmplivre = mid(codigobarras,20,25)
	campo1 = left(codigobarras,4) & mid(cmplivre,1,5)
	campo1 = campo1 & calcdig10(campo1)
	campo1 = mid(campo1,1,5) & "." & mid(campo1,6,5)

	campo2 = mid(cmplivre,6,10)
	campo2 = campo2&calcdig10(campo2)
	campo2 = mid(campo2,1,5) & "." & mid(campo2,6,6)

	campo3 = mid(cmplivre,16,10)
	campo3 = campo3 & calcdig10(campo3)
	campo3 = mid(campo3,1,5) & "." & mid(campo3,6,6)

	campo4 = mid(codigobarras,5,1)

	campo5 = int(mid(codigobarras,6,14))

	if campo5 = 0 then
		campo5 = "000"
	end if

	linhadigitavel_bradesco = campo1 & "&nbsp;&nbsp;" & campo2 & "&nbsp;&nbsp;" & campo3 & "&nbsp;&nbsp;" & campo4 & "&nbsp;&nbsp;" & campo5
'*************************
END FUNCTION
'*************************

'**************************
FUNCTION linhadigitavel_hsbc(codigobarras)
'**************************
	campo1 = left(codigobarras,4) & mid(codigobarras,20,5)
	campo1 = campo1 & calcdig10(campo1)
	campo1 = mid(campo1,1,5) & "." & mid(campo1,6,5)

	campo2 = mid(codigobarras,25,2) & mid(codigobarras,27,8)
	campo2 = campo2 & calcdig10(campo2)
	campo2 = mid(campo2,1,5) & "." & mid(campo2,6,6)

	campo3 = mid(codigobarras,35,5) & mid(codigobarras,40,5)
	campo3 = campo3 & calcdig10(campo3)
	campo3 = mid(campo3,1,5) & "." & mid(campo3,6,6)

	campo4 = mid(codigobarras,5,1)

	campo5 = int(mid(codigobarras,6,14))

	if campo5 = 0 then
		campo5 = "000"
	end if

	linhadigitavel_hsbc = campo1 & "&nbsp;&nbsp;" & campo2 & "&nbsp;&nbsp;" & campo3 & "&nbsp;&nbsp;" & campo4 & "&nbsp;&nbsp;" & campo5
'*************************
END FUNCTION
'*************************

'**************************
FUNCTION linhadigitavel_unibanco(codigobarras)
'**************************
	campo1 = left(codigobarras,4) & mid(codigobarras,20,1) & mid(codigobarras,21,4)
	campo1 = campo1 & calcdig10(campo1)
	campo1 = mid(campo1,1,5) & "." & mid(campo1,6,5)

	campo2 = mid(codigobarras,25,3) & "00" & mid(codigobarras,30,5)
	campo2 = campo2 & calcdig10(campo2)
	campo2 = mid(campo2,1,5) & "." & mid(campo2,6,6)

	campo3 = mid(codigobarras,35,10)
	campo3 = campo3 & calcdig10(campo3)
	campo3 = mid(campo3,1,5) & "." & mid(campo3,6,6)

	campo4 = mid(codigobarras,5,1)

	campo5 = int(mid(codigobarras,6,14))

	if campo5 = 0 then
		campo5 = "000"
	end if

	linhadigitavel_unibanco = campo1 & "&nbsp;&nbsp;" & campo2 & "&nbsp;&nbsp;" & campo3 & "&nbsp;&nbsp;" & campo4 & "&nbsp;&nbsp;" & campo5
'*************************
END FUNCTION
'*************************

'**************************
FUNCTION linhadigitavel_panamericano(codigobarras)
'**************************
	campo1 = left(codigobarras,4) & mid(codigobarras,20,1) & mid(codigobarras,21,4)
	campo1 = campo1 & calcdig10(campo1)
	campo1 = mid(campo1,1,5) & "." & mid(campo1,6,5)

	campo2 = mid(codigobarras,25,3) & "00" & mid(codigobarras,30,5)
	campo2 = campo2 & calcdig10(campo2)
	campo2 = mid(campo2,1,5) & "." & mid(campo2,6,6)

	campo3 = mid(codigobarras,35,10)
	campo3 = campo3 & calcdig10(campo3)
	campo3 = mid(campo3,1,5) & "." & mid(campo3,6,6)

	campo4 = mid(codigobarras,5,1)

	campo5 = int(mid(codigobarras,6,14))

	if campo5 = 0 then
		campo5 = "000"
	end if

	linhadigitavel_panamericano = campo1 & "&nbsp;&nbsp;" & campo2 & "&nbsp;&nbsp;" & campo3 & "&nbsp;&nbsp;" & campo4 & "&nbsp;&nbsp;" & campo5
'*************************
END FUNCTION
'*************************

'**************************
FUNCTION linhadigitavel_real(codigobarras)
'**************************
	cmplivre = mid(codigobarras,20,25)
	campo1 = left(codigobarras,4) & mid(cmplivre,1,5)
	campo1 = campo1 & calcdig10(campo1)
	campo1 = mid(campo1,1,5) & "." & mid(campo1,6,5)

	campo2 = mid(cmplivre,6,10)
	campo2 = campo2&calcdig10(campo2)
	campo2 = mid(campo2,1,5) & "." & mid(campo2,6,6)

	campo3 = mid(cmplivre,16,10)
	campo3 = campo3 & calcdig10(campo3)
	campo3 = mid(campo3,1,5) & "." & mid(campo3,6,6)

	campo4 = mid(codigobarras,5,1)

	campo5 = int(mid(codigobarras,6,14))

	if campo5 = 0 then
		campo5 = "000"
	end if

	linhadigitavel_real = campo1 & "&nbsp;&nbsp;" & campo2 & "&nbsp;&nbsp;" & campo3 & "&nbsp;&nbsp;" & campo4 & "&nbsp;&nbsp;" & campo5
'*************************
END FUNCTION
'*************************

'************************************
FUNCTION CALCDIG10_BRTelecom(pCadeia)
'************************************
	
	'****************************************************************************
	'Calcula o DAC (Digito de auto-conferencia) para codigo CNAB de 44 posicoes
	
	'pCadeia é uma sequencia de 11 digitos, ou seja, um codigo de 44 caracteres
	'possui 4 bloco de 11 digitos, que sao passados para funcao indivualmente.
	'O retorno dessa funcao é o DAC de cada bloco
	'****************************************************************************
	
	'****************************************************************************
	'Calcula o DVG (Digito de verificação geral) para codigo FEBRABAN de 44 posicoes
	
	'O dígito verificador serve para a leitura de cada dígito do Código de Barras e a verificação da sua consistência, fechando o conjunto da informação.

	'Parâmetros:
	'	pCadeia = Bloco de 43 posicoes (44 - 1 [Espaço para o DG a ser gerado])
	
	'Para cálculo do dígito verificador, que deverá constar na quarta posição do Código de Barras, deverá se feita a seguinte montagem:
    '     1.Definir uma área auxiliar de 43 posições subdividida em dois campos. O primeiro de três posições deverá conter, o identificador do produto, identificação do segmento e identificador do valor efetivo ou referência. O segundo campo deverá conter as 40 posições restantes;
    '     2.Calcular o módulo 10, conforme acima, das 43 posições;
    '     3.Montar o campo para impressão no Código de Barras, com as três primeiras posições, o DAC já calculado, e as 40 posições restantes;
    '     4.A representação numérica do Código de Barras, deverá ser montada após o cálculo do dígito verificador.

	'****************************************************************************
	
	pos = 0
	dig_seq = 1
	cadeia_digitos = ""
	length = len(pCadeia) - 1
	
	Do while (pos <= length)
	
		if dig_seq = 1 then dig_seq = 2 else dig_seq = 1 end if
		
		cadeia_digitos = cstr((mid(pCadeia, len(pCadeia)-pos, 1) * dig_seq)) & cadeia_digitos
		pos = pos + 1
	Loop

	soma_cadeia_digitos = 0	
	for i = 1 to len(cadeia_digitos)
		soma_cadeia_digitos = soma_cadeia_digitos + mid(cadeia_digitos,i,1)
	next

	resto = (soma_cadeia_digitos mod 10)

	if resto = 0 then resto = 10
	
	DAC = 10 - resto
	
	CALCDIG10_BRTelecom = DAC 
'*************************
END FUNCTION
'*************************

'**************************
FUNCTION CALCDIG10(cadeia)
'**************************
	mult = (len(cadeia) mod 2) 
	mult = mult + 1
	total = 0
	for pos = 1 to len(cadeia)
		res = mid(cadeia, pos, 1) * mult
		if res > 9 then
			res = int(res/10) + (res mod 10)
		end if
		total = total + res
		if mult = 2 then
			mult = 1
		else
			mult = 2
		end if
	next
	total = ((10-(total mod 10)) mod 10 )
	CALCDIG10 = total
'*************************
END FUNCTION
'*************************

'**************************
FUNCTION CALCDIG11(cadeia,limitesup,lflag)
'**************************
	mult = 1 + (len(cadeia) mod (limitesup-1))
	if mult = 1 then
		mult = limitesup
	end if
	total = 0
	for pos = 1 to len(cadeia)
		total = total + (mid(cadeia,pos,1) * mult)
		mult = mult-1
		if mult = 1 then
			mult = limitesup
		end if
	Next
	nresto = (total mod 11)
	if lflag = 1 then
		calcdig11 = nresto
	else
		if nresto = 0 or nresto = 1 or nresto = 10 then
			ndig = 1
		else
			ndig = 11 - nresto	
		end if
		calcdig11 = ndig
	end if

'*************************
END FUNCTION
'*************************

'**************************
FUNCTION CALCDIG11_Santander(cadeia,limitesup,lflag)
'**************************
	mult = 1 + (len(cadeia) mod (limitesup-1))
	if mult = 1 then
		mult = limitesup
	end if
	total = 0
	for pos = 1 to len(cadeia)
		total = total + (mid(cadeia,pos,1) * mult)
		mult = mult-1
		if mult = 1 then
			mult = limitesup
		end if
	Next
	nresto = (total mod 11)
	if nresto = 0 then
		ndig = 1
	elseif nresto = 1 then
		ndig = 0
	else
		ndig = 11 - nresto	
	end if
	CALCDIG11_Santander = ndig

'*************************
END FUNCTION
'*************************

'**************************
FUNCTION CALCDIG11_Bradesco(cadeia,limitesup,lflag)
'**************************
	mult = 1 + (len(cadeia) mod (limitesup-1))
	if mult = 1 then
		mult = limitesup
	end if
	total = 0
	for pos = 1 to len(cadeia)
		total = total + (mid(cadeia,pos,1) * mult)
		mult = mult-1
		if mult = 1 then
			mult = limitesup
		end if
	Next
	nresto = (total mod 11)
	if lflag = 1 then
		CALCDIG11_Bradesco = nresto
	else
		if nresto = 1 then
			ndig = "P"
		elseif nresto = 0 then
			ndig = 0
		else
			ndig = 11 - nresto	
		end if
		CALCDIG11_Bradesco = ndig
	end if

'*************************
END FUNCTION
'*************************

'**************************
FUNCTION CALCDIG11_HSBC(cadeia,limitesup,lflag)
'**************************
	mult = 1 + (len(cadeia) mod (limitesup-1))
	if mult = 1 then
		mult = limitesup
	end if
	total = 0
	for pos = 1 to len(cadeia)
		total = total + (mid(cadeia,pos,1) * mult)
		mult = mult-1
		if mult = 1 then
			mult = limitesup
		end if
	Next
	nresto = (total mod 11)
	if nresto = 0 or nresto = 1 or nresto = 10 then
		ndig = 0
	else
		ndig = 11 - nresto	
	end if
	CALCDIG11_HSBC = ndig

'*************************
END FUNCTION
'*************************

'**************************
'Calcula o SUPERDIGITO do UNIBANCO
'**************************
FUNCTION SUPERDIGITO(CNOSSO)  
	numerodedigitos = len(cnosso)

	ATAB(0) = 8
	ATAB(1) = 7
	ATAB(2) = 6
	ATAB(3) = 5
	ATAB(4) = 4
	ATAB(5) = 3
	ATAB(6) = 2
	ATAB(7) = 9
	ATAB(8) = 8
	ATAB(9) = 7
	ATAB(10) = 6
	ATAB(11) = 5
	ATAB(12) = 4
	ATAB(13) = 3
	ATAB(14) = 2 
	NSOMA = 0
	NUNIDADE = 0
	NDIGITO = 0
	NCONTA = numerodedigitos
	while NCONTA >= 1
		NUNIDADE1 = MID(CNOSSO,NCONTA,1)
		NUNIDADE = MID(CNOSSO,NCONTA,1) * ATAB(NCONTA)
		NSOMA = NSOMA + NUNIDADE
		NCONTA = NCONTA - 1
	wend
	digito = (NSOMA*10) mod 11
	if digito  = 0 or digito  = 10 then
		digito = 0
	else
	end if
	SUPERDIGITO = digito

'*************************
END FUNCTION
'*************************

'CALCULA O DIGITO VERIFICADOR DO CÓDIGO DE BARRAS E É BARRA
'**************************
FUNCTION CALCNUMB(CNOSSO)
'**************************
	numerodedigitos = len(cnosso)

	IF numerodedigitos = "44" THEN
		ATAB(0)=6
		atab(1)=5
		atab(2)=4
		ATAB(3)=3
		ATAB(4)=2
		ATAB(5)=9
		ATAB(6)=8
		ATAB(7)=7
		ATAB(8)=6
		ATAB(9)=5
		ATAB(10)=4
		ATAB(11)=3
		ATAB(12)=2
		ATAB(13)=9
		ATAB(14)=8
		ATAB(15)=7
		ATAB(16)=6
		ATAB(17)=5
		ATAB(18)=4
		ATAB(19)=3
		ATAB(20)=2
		ATAB(21)=9
		ATAB(22)=8
		ATAB(23)=7
		ATAB(24)=6
		ATAB(25)=5
		ATAB(26)=4
		ATAB(27)=3
		ATAB(28)=2
		ATAB(29)=9
		ATAB(30)=8
		ATAB(31)=7
		ATAB(32)=6
		ATAB(33)=5
		ATAB(34)=4
		ATAB(35)=3
		ATAB(36)=2
		ATAB(37)=9
		ATAB(38)=8
		ATAB(39)=7
		ATAB(40)=6
		ATAB(41)=5
		ATAB(42)=4
		ATAB(43)=3
		ATAB(44)=2
	ELSE
		ATAB(0)=5
		atab(1)=4
		atab(2)=3
		ATAB(3)=2
		ATAB(4)=9
		ATAB(5)=8
		ATAB(6)=7
		ATAB(7)=6
		ATAB(8)=5
		ATAB(9)=4
		ATAB(10)=3
		ATAB(11)=2
		ATAB(12)=9
		ATAB(13)=8
		ATAB(14)=7
		ATAB(15)=6
		ATAB(16)=5
		ATAB(17)=4
		ATAB(18)=3
		ATAB(19)=2
		ATAB(20)=9
		ATAB(21)=8
		ATAB(22)=7
		ATAB(23)=6
		ATAB(24)=5
		ATAB(25)=4
		ATAB(26)=3
		ATAB(27)=2
		ATAB(28)=9
		ATAB(29)=8
		ATAB(30)=7
		ATAB(31)=6
		ATAB(32)=5
		ATAB(33)=4
		ATAB(34)=3
		ATAB(35)=2
		ATAB(36)=9
		ATAB(37)=8
		ATAB(38)=7
		ATAB(39)=6
		ATAB(40)=5
		ATAB(41)=4
		ATAB(42)=3
		ATAB(43)=2
	END IF
	NSOMA = 0
	NUNIDADE = 0
	NDIGITO = 0
	numerodedigitos = len(cnosso)
	NCONTA = numerodedigitos

	while NCONTA >= 1
		NUNIDADE = MID(CNOSSO,NCONTA,1) * ATAB(NCONTA)    
		NSOMA = NSOMA + NUNIDADE
		NCONTA = NCONTA - 1
	wend
	digito = (NSOMA*10) mod 11
	if digito = 0 or digito = 10 then
		digito = 0
	else
	end if
	CALCNUMB = digito
'*************************
END FUNCTION
'*************************

'**************************
FUNCTION fatorvencimento(vencimento)
'**************************

	if len(vencimento) < 8 then
		fatorvencimento = "0000"
	else
		fatorvencimento = datevalue("" & vencimento & "") - datevalue("1997/10/07")
	end if

'*************************
END FUNCTION
'*************************

'**************************
FUNCTION campolivre(CEDENTE,nossonumero,banco)
'**************************
	campolivre = cedente & nossonumero & "00" & banco
	campolivre = campolivre & calcdig10(campolivre)

	do while true
		cauxiliar = calcdig11(campolivre,7,1)
		if cauxiliar = 0 then
			exit do
		elseif cauxiliar = 1 then
			if right(campolivre,1) = 9 then
				campolivre = mid(campolivre,1,len(campolivre)-1)
				campolivre = campolivre & "0"
			else
				ultimo = right(campolivre,1) + 1
				campolivre = mid(campolivre,1,len(campolivre)-1)
				campolivre = campolivre & ultimo
			end if
		else
			cauxiliar = 11 - cauxiliar
			exit do
		end if	
	loop
	campolivre = campolivre & cauxiliar

'*************************
END FUNCTION
'*************************

'**************************
FUNCTION codbar_caixa(banco,moeda,vencimento,valor,cedente,agencia,nossonumero)
'**************************
	strcodbar = banco & moeda & vencimento & valor & cedente & agencia & "87" & nossonumero
	d3 = calcdig11(strcodbar,9,0)
	strcodbar = banco & moeda & d3 & vencimento & valor & cedente & agencia & "87" & nossonumero
	codbar_caixa = strcodbar
'*************************
END FUNCTION
'*************************

'**************************
FUNCTION codbar_banespa(banco,moeda,vencimento,valor,cedente,nossonumero)
'**************************
	campolivre1 = campolivre(cedente,nossonumero,banco)
	strcodbar = banco & moeda & vencimento & valor & campolivre1
	d3 = calcdig11(strcodbar,9,0)
	strcodbar = banco & moeda & d3 & vencimento & valor & campolivre1
	codbar_banespa = strcodbar
'*************************
END FUNCTION
'*************************

'**************************
FUNCTION codbar_santander(banco,moeda,vencimento,valor,cedente,nossonumero)
'**************************
	strcodbar = banco & moeda & vencimento & valor & "9" & cedente & nossonumero & "0102"
	d3 = calcdig11(strcodbar,9,0)
	codbar_santander = banco & moeda & d3 & vencimento & valor & "9" & cedente & nossonumero & "0102"
'*************************
END FUNCTION
'*************************

'**************************
FUNCTION codbar_itau(banco,moeda,vencimento,valor,carteira,nossonumero,dvnossonumero,agencia,conta,dvagconta)
'**************************
	strcodbar = banco & moeda & vencimento & valor & carteira & nossonumero & dvnossonumero & agencia & conta & dvagconta & "000"
	dv3 = calcdig11(strcodbar,9,0)
	codbar_itau = banco & moeda & dv3 & vencimento & valor & carteira & nossonumero & dvnossonumero & agencia & conta & dvagconta & "000"
'*************************
END FUNCTION
'*************************

'**************************
FUNCTION codbar_hsbc(banco,moeda,vencimento,valor,numcliente,nossonumero,datavencimento)
'**************************
	data_juliano = Right("00" & Datediff("d","01/01/" & Year(datavencimento),datavencimento),3) & Right(Year(datavencimento),1)
	strcodbar = banco & moeda & vencimento & valor & numcliente & nossonumero & data_juliano & "2"
	dv3 = calcdig11(strcodbar,9,0)
	codbar_hsbc = banco & moeda & dv3 & vencimento & valor & numcliente & nossonumero & data_juliano & "2"
'*************************
END FUNCTION
'*************************

'**************************
FUNCTION codbar_unibanco(banco,moeda,vencimento,valor,numcliente,nossonumero)
'**************************
	dv_nossonumero = SUPERDIGITO(nossonumero)
	strcodbar = banco & moeda & vencimento & valor & "5" & numcliente & "00" & nossonumero & dv_nossonumero
	dv3 = calcdig11(strcodbar,9,0)
	'dv3 = CALCNUMB(strcodbar)
	codbar_unibanco = banco & moeda & dv3 & vencimento & valor & "5" & numcliente & "00" & nossonumero & dv_nossonumero
'*************************
END FUNCTION
'*************************

'**************************
FUNCTION codbar_bradesco(banco,moeda,vencimento,valor,agencia,carteira,nossonumero,contacorrente)
'**************************
	strcodbar = banco & moeda & vencimento & valor & agencia & carteira & nossonumero & contacorrente & "0"
	dv3 = calcdig11(strcodbar,9,0)
	'dv3 = CALCNUMB(strcodbar)
	codbar_bradesco = banco & moeda & dv3 & vencimento & valor & agencia & carteira & nossonumero & contacorrente & "0"
'*************************
END FUNCTION
'*************************

'**************************
FUNCTION codbar_panamericano(banco,moeda,vencimento,valor,agencia,carteira,operacao,nossonumero)
'**************************
	strcodbar = banco & moeda & vencimento & valor & agencia & carteira & operacao & nossonumero
	dv3 = calcdig11(strcodbar,9,0)
	codbar_panamericano = banco & moeda & dv3 & vencimento & valor & agencia & carteira & operacao & nossonumero
'*************************
END FUNCTION
'*************************

'**************************
FUNCTION codbar_real(banco,moeda,agencia,contacorrente,dvnossonumero,nossonumero,vencimento,valor)
'**************************
	strcodbar = banco & moeda & vencimento & valor & agencia & contacorrente & dvnossonumero & nossonumero
	dv3 = calcdig11(strcodbar,9,0)
	codbar_real = banco & moeda & dv3 & vencimento & valor & agencia & contacorrente & dvnossonumero & nossonumero
'*************************
END FUNCTION
'*************************

'**************************
FUNCTION CALCNUMB2(CNOSSO)
'**************************
	atab(0)=7
	atab(1)=7
	ATAB(2)=3
	ATAB(3)=1
	ATAB(4)=9
	ATAB(5)=7
	ATAB(6)=3
	ATAB(7)=1
	ATAB(8)=9
	ATAB(9)=7
	ATAB(10)=3
	NSOMA = 0
	NUNIDADE = 0
	NDIGITO = 0

	FOR NCONTA = 1 TO 10
		'NUNIDADE = MID(CNOSSO,NCONTA,1)
		NUNIDADE = MID(CNOSSO,NCONTA,1) * ATAB(NCONTA)
		NUNIDADE = RIGHT(NUNIDADE,1)
		NSOMA = NSOMA + NUNIDADE
	NEXT

	nsoma = right(nsoma,1)
	if nsoma = 0 then
		ndigito = 0
	else
		ndigito = 10 - nsoma
	end if
	CALCNUMB2 = ndigito
'*************************
END FUNCTION
'*************************

'Desenho da barra
'**************************
Sub WBarCode( Valor )
'**************************

	Dim f, f1, f2, i
	Dim texto
	Const fino = 1
	Const largo = 3
	Const altura = 50
	Dim BarCodes(99)

	if isempty(BarCodes(0)) then
		BarCodes(0) = "00110"
		BarCodes(1) = "10001"
		BarCodes(2) = "01001"
		BarCodes(3) = "11000"
		BarCodes(4) = "00101"
		BarCodes(5) = "10100"
		BarCodes(6) = "01100"
		BarCodes(7) = "00011"
		BarCodes(8) = "10010"
		BarCodes(9) = "01010"
		for f1 = 9 to 0 step -1
			for f2 = 9 to 0 Step -1
				f = f1 * 10 + f2
				texto = ""
				for i = 1 To 5
					texto = texto & mid(BarCodes(f1), i, 1) + mid(BarCodes(f2), i, 1)
				next
				BarCodes(f) = texto
			next
		next
	end if
	' Guarda inicial
	%>
	<img src=images/2.gif width=<%=fino%> height=<%=altura%> border=0><img 
	src=images/1.gif width=<%=fino%> height=<%=altura%> border=0><img 
	src=images/2.gif width=<%=fino%> height=<%=altura%> border=0><img 
	src=images/1.gif width=<%=fino%> height=<%=altura%> border=0><img 
	<%
	texto = valor
	if len( texto ) mod 2 <> 0 then
		texto = "0" & texto
	end if


	' Draw dos dados
	do while len(texto) > 0
		i = cint( left( texto, 2) )
		texto = right( texto, len( texto ) - 2)
		f = BarCodes(i)
		for i = 1 to 10 step 2
			if mid(f, i, 1) = "0" then
				f1 = fino
			else
				f1 = largo
			end if
			%>
			src=images/2.gif width=<%=f1%> height=<%=altura%> border=0><img 
			<%
			if mid(f, i + 1, 1) = "0" Then
				f2 = fino
			else
				f2 = largo
			end if
			%>
			src=images/1.gif width=<%=f2%> height=<%=altura%> border=0><img 
			<%
		next
	loop

	' Draw guarda final
	%>
	src=images/2.gif width=<%=largo%> height=<%=altura%> border=0><img 
	src=images/1.gif width=<%=fino%> height=<%=altura%> border=0><img 
	src=images/2.gif width=<%=1%> height=<%=altura%> border=0>
	<%
'**************************
end sub
'**************************

%>

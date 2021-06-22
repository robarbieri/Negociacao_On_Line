	function fnIsDate(str)
	{
	try
		{
		//Se String vier vazia retorna
		if (str=='')
			{
			return(true);
			}
		//Valida Expressão dd/mm/aaaa
		var dtRegExp = /^(\d{10}|\d{2}\/\d{2}\/\d{4})$/;
		if (dtRegExp.test(str))
			{
			var strData = new String(str);
			var arrData = strData.split('/');
			var dia = new Number(arrData[0]);
			var mes = new Number(arrData[1]);
			var ano = new Number(arrData[2]);
			//Verifica se dia e mes existe
			if (dia < 1 || dia > 31 || mes < 1 || mes > 12)
				{
				return(false);
				}
			//Verifica se dia existe no mes	
			if ((mes == 4 || mes == 6 || mes == 9 || mes == 11) && (dia == 31))
				{
				return(false);
				}
			//Se Fevereiro	
			if (mes == 2)
				{
				//Se for Ano-Bissexto
				if (ano % 4 == 0)
					{
					if (dia > 29)
						{
						return(false);
						}
					}
				//Se não for Ano-Bissexto	
				else
					{
					if (dia > 28)
						{
						return(false);
						}				
					}				
				}
			return(true);
			}
		else
			{
			return(false);
			}
		}
	catch(e)
		{
		return(false);		
		}
	}
	
	
	function FormatarCampoData()
		{
		try
			{
			with(window.event.srcElement)
				{
				maxLength = 10			
	/*			if((event.keyCode < 48 || event.keyCode > 57) && (event.keyCode < 96 || event.keyCode > 105) && event.keyCode != 46 && event.keyCode != 8 && event.keyCode != 16 && event.keyCode != 37 && event.keyCode != 9)*/
				if((event.keyCode < 48 || event.keyCode > 57) && (event.keyCode < 96 || event.keyCode > 105) && event.keyCode != 46 && event.keyCode != 8 && event.keyCode != 16 && event.keyCode != 37 && event.keyCode != 9)
					{
					event.returnValue = false;		
					}					
				if((value.length == 2 ||value.length == 5) && event.keyCode != 8)
					{
					value += "/";
					}
				}		
			}	
		catch(e)
			{
			event.returnValue = false;
			}
		}

	function FormatarCampoEMail()
		{
		with(window.event.srcElement)
			{
			maxLength = 60
			if ((event.keyCode < 48 && event.keyCode != 45 && event.keyCode != 46) || (event.keyCode > 122 && event.keyCode != 190))
				{				
				event.returnValue = false;		
				}
			if(event.keyCode > 90 && event.keyCode < 97 && event.keyCode != 95)
				{				
				event.returnValue = false;		
				}
			if(event.keyCode > 57 && event.keyCode < 64)
				{				
				event.returnValue = false;		
				}
			}
		}

	function FormatarCampoFONE()
		{
		try
			{
			with(window.event.srcElement)
				{
				maxLength = 9
				if(event.keyCode < 48 || event.keyCode > 57)
					{
					event.returnValue = false;		
					}
				if(value.length == 3 && event.keyCode != 8)
					{
					value += '-';
					}
				if(value.length == 8 && event.keyCode != 8)
					{
					value = value.replace('-','');
					value = value.substr(0, 4) + '-' + value.substr(4)
					}
				}		
			}	
		catch(e)
			{
			event.returnValue = false;
			}
		}

	function FormatarCampoNumero()
		{
		try
			{
			with(window.event.srcElement)
				{
				if((event.keyCode < 48 || event.keyCode > 57) && (event.keyCode < 96 || event.keyCode > 105) && event.keyCode != 46 && event.keyCode != 8 && event.keyCode != 16 && event.keyCode != 37 && event.keyCode != 9)
					{
					event.returnValue = false;		
					}					
				}		
			}	
		catch(e)
			{
			event.returnValue = false;
			}
		}

    function FormataValor(campo,tammax,teclapres) {
	    var tecla = teclapres.keyCode;
    //	vr = document.form1.item(campo).value;
	    vr = document.form1.elements[campo].value
	    vr = vr.replace( "/", "" );
	    vr = vr.replace( "/", "" );
	    vr = vr.replace( ",", "" );
	    vr = vr.replace( ".", "" );
	    vr = vr.replace( ".", "" );
	    vr = vr.replace( ".", "" );
	    vr = vr.replace( ".", "" );
	    tam = vr.length;

	    if (tam < tammax && tecla != 8){ tam = vr.length + 1 ; }

	    if (tecla == 8 ){	tam = tam - 1 ; }
    		
	    if ( tecla == 8 || tecla >= 48 && tecla <= 57 || tecla >= 96 && tecla <= 105 ){
		    if ( tam <= 2 ){ 
	 		    document.form1.elements[campo].value = vr ; }
	 	    if ( (tam > 2) && (tam <= 5) ){
	 		    document.form1.elements[campo].value = vr.substr( 0, tam - 2 ) + ',' + vr.substr( tam - 2, tam ) ; }
	 	    if ( (tam >= 6) && (tam <= 8) ){
	 		    document.form1.elements[campo].value = vr.substr( 0, tam - 5 ) + '.' + vr.substr( tam - 5, 3 ) + ',' + vr.substr( tam - 2, tam ) ; }
	 	    if ( (tam >= 9) && (tam <= 11) ){
	 		    document.form1.elements[campo].value = vr.substr( 0, tam - 8 ) + '.' + vr.substr( tam - 8, 3 ) + '.' + vr.substr( tam - 5, 3 ) + ',' + vr.substr( tam - 2, tam ) ; }
	 	    if ( (tam >= 12) && (tam <= 14) ){
	 		    document.form1.elements[campo].value = vr.substr( 0, tam - 11 ) + '.' + vr.substr( tam - 11, 3 ) + '.' + vr.substr( tam - 8, 3 ) + '.' + vr.substr( tam - 5, 3 ) + ',' + vr.substr( tam - 2, tam ) ; }
	 	    if ( (tam >= 15) && (tam <= 17) ){
	 		    document.form1.elements[campo].value = vr.substr( 0, tam - 14 ) + '.' + vr.substr( tam - 14, 3 ) + '.' + vr.substr( tam - 11, 3 ) + '.' + vr.substr( tam - 8, 3 ) + '.' + vr.substr( tam - 5, 3 ) + ',' + vr.substr( tam - 2, tam ) ;}
	    }	
    }

    function ReFormataValor(campo,tammax,teclapres) {
	    var tecla = teclapres.keyCode;
    //	vr = document.form1.item(campo).value;
	    vr = document.form1.elements[campo].value;
	    vr = vr.replace( "/", "" );
	    vr = vr.replace( "/", "" );
	    vr = vr.replace( ",", "" );
	    vr = vr.replace( ".", "" );
	    vr = vr.replace( ".", "" );
	    vr = vr.replace( ".", "" );
	    vr = vr.replace( ".", "" );
    	
	    tam = vr.length;

	    vr2 = vr.substr(tam - 1, tam);
    	
	    if (vr2 != '.' && vr2 != ',' && vr2 != 0 && vr2 != 1 && vr2 != 2 && vr2 != 3 && vr2 != 4 && vr2 != 5 && vr2 != 6 && vr2 != 7 && vr2 != 8 && vr2 != 9){
		    vr = vr.substr(0, tam - 1);
	    }
    	
	    tam = vr.length;

	    if ( tam <= 2 ){ 
		    document.form1.elements[campo].value = vr ; }
	    if ( (tam > 2) && (tam <= 5) ){
		    document.form1.elements[campo].value = vr.substr( 0, tam - 2 ) + ',' + vr.substr( tam - 2, tam ) ; }
	    if ( (tam >= 6) && (tam <= 8) ){
		    document.form1.elements[campo].value = vr.substr( 0, tam - 5 ) + '.' + vr.substr( tam - 5, 3 ) + ',' + vr.substr( tam - 2, tam ) ; }
	    if ( (tam >= 9) && (tam <= 11) ){
		    document.form1.elements[campo].value = vr.substr( 0, tam - 8 ) + '.' + vr.substr( tam - 8, 3 ) + '.' + vr.substr( tam - 5, 3 ) + ',' + vr.substr( tam - 2, tam ) ; }
	    if ( (tam >= 12) && (tam <= 14) ){
		    document.form1.elements[campo].value = vr.substr( 0, tam - 11 ) + '.' + vr.substr( tam - 11, 3 ) + '.' + vr.substr( tam - 8, 3 ) + '.' + vr.substr( tam - 5, 3 ) + ',' + vr.substr( tam - 2, tam ) ; }
	    if ( (tam >= 15) && (tam <= 17) ){
		    document.form1.elements[campo].value = vr.substr( 0, tam - 14 ) + '.' + vr.substr( tam - 14, 3 ) + '.' + vr.substr( tam - 11, 3 ) + '.' + vr.substr( tam - 8, 3 ) + '.' + vr.substr( tam - 5, 3 ) + ',' + vr.substr( tam - 2, tam ) ;}
    }
	
	function validaCampos(){
	    var objDateNow = new Date();
	    if(document.form1.txtVencto.value==""){
	        var objDateVencto = new Date();
		    var objDateVencto = objDateVencto.add("d",5);
			document.form1.txtVencto.value = Right("00" + parseInt(objDateVencto.getDate()),2) + "/" + Right("00" + parseInt(parseInt(objDateVencto.getMonth()) + parseInt(1)),2) + "/" + objDateVencto.getFullYear();
		}
		if(! fnIsDate(document.form1.txtVencto.value)){
			alert("Vencimento da parcela deve estar no formato DD/MM/AAAA.");
			document.form1.txtVencto.select();
			return false;
		}
		 
		var objDate = new Date();
		var objDateAdd = new Date();
		var objDateAdd = objDateAdd.add("d",15);
		var datToday = (objDateAdd.getFullYear() + "" + Right("00" + parseInt(parseInt(objDateAdd.getMonth()) + parseInt(1)),2) + "" + Right("00" + parseInt(objDateAdd.getDate()),2));
		var datNow = (objDate.getFullYear() + "" + Right("00" + parseInt(parseInt(objDate.getMonth()) + parseInt(1)),2) + "" + Right("00" + parseInt(objDate.getDate()),2));
		var arrDat_Prioridade = document.form1.txtVencto.value.split("/");
		var datDat_Prioridade = (arrDat_Prioridade[2] + "" + Right("00" + parseInt(arrDat_Prioridade[1]),2) + "" + Right("00" + parseInt(arrDat_Prioridade[0]),2));
        
//        alert(datDat_Prioridade + " " + datToday + " " + datNow);
        if(parseInt(datDat_Prioridade) <= parseInt(datNow)){
			alert("Vencimento da parcela deve ser maior que hoje.");
			document.form1.txtVencto.select();
			return false;
		}
		
		if(parseInt(datDat_Prioridade) > parseInt(datToday)){
			alert("O vencimento da parcela deve ser menor que hoje + 15 dias.");
			document.form1.txtVencto.select();
			return false;
		}
		
		if(parseInt(document.form1.txtValor.value) < 50){
		    alert("O valor mínimo da entrada é de R$50,00.");
			document.form1.txtValor.select();
			return false;
		}
		
		//alert(parseInt(Replace(document.form1.txtValor.value,".","")) + " " + parseInt(document.form1.Atualizado.value))
		if(parseInt(Replace(document.form1.txtValor.value,".","")) > parseInt(document.form1.Atualizado.value)){
			alert("O valor da entrada deve ser menor que o valor do Acordo.");
			document.form1.txtVencto.select();
			return false;
		}
		
		if(document.form1.txtValor.value==""){
			document.form1.txtValor.value='0';
		}
		
		if(Left(document.form1.cboCondicao.value,2) == "01"){
			document.form1.txtValor.value='0';
		}
		
		if(document.form1.cboCondicao.value=="Selecione..."){
			alert("Selecione o parcelamento.");
			document.form1.cboCondicao.select();
			return false;
		}
		return true;
	}
	
	function validaCamposAtz(){
		if(document.form1.txtEMail.value=""){
			alert("Vencimento da parcela deve estar no formato DD/MM/AAAA.");
			document.form1.txtVencto.select();
			return false;
		}
	return true;
	}
	
	function verificaForm()
	{
		if(! validaCampos())
		{
			return false;
	    }
	    else
	    {
//	    alert(document.form1.Principal.value + " " +
//		document.form1.valCorrecao.value + " " +
//        document.form1.PercentHonor.value + " " +
//        document.form1.cboCondicao.value + " " +
//        document.form1.txtValor.value);
		calcularAcordo();
		return true;
	    }
	}
	
//	function calcularAcordo(){
//		document.getElementById("lblParcela").innerHTML = "Calculando.......Aguarde.";
//		    inNegociacao.Calcular(document.form1.Principal.value,
//		                          document.form1.valCorrecao.value,
//		                          document.form1.PercentHonor.value,
//		                          Replace(document.form1.cboCondicao.value,"x",""),
//		                          Replace(document.form1.txtValor.value,".",""),
//		                          document.form1.txtVencto.value,
//		                          document.form1.txtidLogin.value,
//		                          document.form1.txtContratoOS.value,
//		    calcularAcordo_CallBack);
//	}
	
//	function calcularAcordo_CallBack(response)
//	    {
//		    document.getElementById("lblParcela").innerHTML = response.value;
//	    }
	
    function FormatBrasil(Valor)
    {
	    var strTmp = Valor;
	    strTmp = Replace(strTmp,",","@");
	    strTmp = Replace(strTmp,".",",");
	    strTmp = Replace(strTmp,"@",".");
	    return strTmp;
    }
    
    function Replace(Expression, Find, Replace)
    {
	    var temp = Expression;
	    var a = 0;

	    for (var i = 0; i < Expression.length; i++) 
	    {
		    a = temp.indexOf(Find);
		    if (a == -1){
			    break
			}
		    else{
			    temp = temp.substring(0, a) + Replace + temp.substring((a + Find.length));
				}
	    }
	    return temp;
    }
    
    Date.prototype.add = function (sInterval, iNum){
      var dTemp = this;
      if (!sInterval || iNum == 0) return dTemp;
      switch (sInterval.toLowerCase()){
        case "ms":
          dTemp.setMilliseconds(dTemp.getMilliseconds() + iNum);
          break;
        case "s":
          dTemp.setSeconds(dTemp.getSeconds() + iNum);
          break;
        case "mi":
          dTemp.setMinutes(dTemp.getMinutes() + iNum);
          break;
        case "h":
          dTemp.setHours(dTemp.getHours() + iNum);
          break;
        case "d":
          dTemp.setDate(dTemp.getDate() + iNum);
          break;
        case "mo":
          dTemp.setMonth(dTemp.getMonth() + iNum);
          break;
        case "y":
          dTemp.setFullYear(dTemp.getFullYear() + iNum);
          break;
      }
      return dTemp;
    }
  
    
    function Left(str, n){
	    if (n <= 0)
	        return "";
	    else if (n > String(str).length)
	        return str;
	    else
	        return String(str).substring(0,n);
    }

    
    function Right(str, n){
        if (n <= 0)
           return "";
        else if (n > String(str).length)
           return str;
        else {
           var iLen = String(str).length;
           return String(str).substring(iLen, iLen - n);
        }
    }   

    function LTrim( value ) {
    	
	    var re = /\s*((\S+\s*)*)/;
	    return value.replace(re, "$1");
    	
    }


    function RTrim( value ) {
    	
	    var re = /((\s*\S+)*)\s*/;
	    return value.replace(re, "$1");
    	
    }


    function trim( value ) {
    	
	    return LTrim(RTrim(value));
    	
    }

    function popupcenter(Url,Name,PosFimX,PosFimY,ScrollBars,Resizable)
    {

      PosIniX=((screen.availWidth/2)-(PosFimX/2));

      PosIniY=((screen.availHeight/2)-(PosFimY/2));

      window.open(Url,Name,'toolbar=0,location=0,directories=0,menubar=0,scrollbars='+ScrollBars+',resizable='+Resizable+',top='+PosIniY+',left='+PosIniX+',width='+PosFimX+',height='+PosFimY+'');

    }
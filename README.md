#include "RWMAKE.CH"   
#include "fivewin.ch"
#include "tcbrowse.ch"
#Include "colors.ch"  
#Include "colors.ch"  
#Include "colors.ch"  
#include "Ap5Mail.ch"

//Autor Paulo H.L.Oliveira
//Rotina de importação E.D.I. GAUSS
//Rede Nacional de Dados Pagina : 2
//Sub-Comissao de E.D.I. Transacao : 601
//Manual Operacional Versao : 03
//PEDIDO MERCADORIA ANCAP/SINDIP Vigencia : 01/11/96
//Emissao : 24/07/02

User Function GAEDIR1()
//alert("GAEDir1")

	lMSErroAuto := .F.
	lMSHelpAuto := .T.

	
	lAutoErrNoFile := .T. 
   	IF MsgYesNo("T E M  C E R T E Z A   D A   I M P O R T A C A O  DE PEDIDOS??? ")      //
		   	Processa({|| okProc2()}) 
	   	  // 	 alert("Fim")     //
		   	
 	ELSE     //
	 	Return    //                                                                                        
	Endif		//	    
                                       
Return

//Inicia Processo de Importação, lendo arquivos da pasta \EDI\RECEBIDOS\

Static Function okProc2()                     

	_cdCliente	:=''
	_cdLoja     :=''
	cIdPrd      :=''
	_nmReduz    :=''
	nPedComp    :=''
	dDtaComp    :=''
	cIdPed      :=''
	dDtaComp	:='' 
	_ufCli		:='' 
	_cVend		:=''
	_cTab		:=''
	_cMun		:=''
	_a1email	:=''
	_a3email	:=''
	_vnomeC		:=''
	_vnomeV		:=''	
	lOk := .T.
	aDados := {}
	nX := 0
   	_count:=0   
	_lidos:=0                                       
 	cDir    	:= 	"C:\PEDIDOSGAUSS\ENVIAR\"  
 	cDirOld		:=	"C:\PEDIDOSGAUSS\ENVIAR\FEITOS\"  
	_aFiles  := Directory( cDir + "*.*" )
	aSort( _aFiles,,, { |x, y| x[1] < y[1] })
    procregua(len(_afiles))
   // alert ("Total arquivos "+ str(len(_aFiles)))  

	if !Empty( _aFiles ) 
		Z:=1 
//		alert(" Z Atual "+ str(Z)) 
	    
		While Z <= (Len(_aFiles)   )               
//			For I := 1 to (Len(_aFiles)   )

				i:=Z
				okSegue()
				Z:= Z+ 1 
 //				alert ("proximo "+ str(Z))
 //			Next
  		Enddo
  		
    Endif  
   
Return    

Static Function okSegue()		
    
    
//	if !Empty( _aFiles )
 //		For I := 1 to (Len(_aFiles)+ 1    )
			_lidos:=_lidos+1
			_carqTxt:=cdir+Trim(_aFiles[i,1]) 
			_carqOld:=cDirOld+Trim(_aFiles[i,1]) 
			hFile 	:= fOpen(_cArqTxt,68) 
 			cEOL    := "CHR(13)+CHR(10)"
			If hFile == -1
		    	MsgAlert("O arquivo de nome "+_cArqTxt+" nao pode ser aberto! Verifique os parametros.","Atencao!")//jb
		    	Return
			Endif                        
	  		importa()      //importa arquivo
	   		fClose(hFile) // libera arquivo
			okOld()       //remove arquivo
   	        incproc ("2. Tratando... "+alltrim(str(I)))
  //	Next 	
	
 	//Endif        
  
Return

Static FUNCTION importa()
	Private cEOL := CHR(13) + CHR(10)
	Private lBuffer := 1000
	Private lFilePos := 1000
	Private lPos := 0
	Private cLine := ""
	Private xCodEmp := "99"
	Private xCodFil     := "01"
	Private cMenNota:=''    
	Private cTextL3 :=''
	Private lMultB :=.F. //controle de multiplos 
	Private lBlqProd :=.F.// controle de bloquei de Produto  
	Private lErroProd :=.F.//controle no erro do código do produto

	lFilePos := FSEEK(hFile, 0, 0)              // POSICIONA PONTEIRO DO ARQUIVO NO PRIMEIRO CARACTER
	cBuffer := SPACE(lBuffer)                   // ALOCA BUFFER
	lRead := FREAD(hFile, cBuffer, lBuffer)     // LE OS PRIMEIROS 1000 CARACTERES DO ARQUIVO
	lPos := AT(cEOL, cBuffer)                   // PROCURA O PRIMEIRO FINAL DE LINHA
    _count := 0   
   	cDoc := GetSxeNum("SC5","C5_NUM")
	aCabPV  := {}   
	_aItens :={}  
	aItemPV := {} 
	_XXROTIN:= ''	
	_it		:= 0 
	_itN	:= 0
	_cIt   	:= STRZERO(0,TamSX3("C6_ITEM")[1]) 
	_cItNok := 0
	_PedEdi:="N" 
	_MNSCLI :="PEDIDO CLIENTE:"   
	_CPEDCLI:=''

	WHILE !(lRead == 0) //.and. _count < 6 

    
    	WHILE (lPos == 0)                               // SE CARACTER DE FINAL DE LINHA NAO FOR ENCONTRADO

	        lBuffer += 1000                             // AUMENTA TAMANHO DO BUFFER

	        cBuffer := SPACE(lBuffer)                   // REALOCA BUFFER
	        lFilePos := FSEEK(hFile, lFilePos, 0)       // REPOSICIONA PONTEIRO DO ARQUIVO
	        lRead := FREAD(hFile, cBuffer, lBuffer)     // LE OS CARACTERES DO ARQUIVO
	        lPos := AT(cEOL, cBuffer)                   // PROCURA O PRIMEIRO FINAL DE LINHA
   		END   
	
    // LEITURA DOS CAMPOS E GRAVACAO DOS DADOS DA TABELA AQUI
 
    	cLine := SUBSTR(cBuffer, 0, lPos)
        
	    if (SUBSTR(cLine, 1,3))=='PP1'
        
        		nItemDoc	:=	(SUBSTR(cLine, 4,4))
        		nCasDec		:=	(SUBSTR(cLine, 8,1))
        		nPedComp	:=	(SUBSTR(cLine, 9,12))
        		dDtaComp	:=	(SUBSTR(cLine, 21,6))
        		cIdPedCo	:=	(SUBSTR(cLine, 27,12))
        		dVigPedC	:=	(SUBSTR(cLine, 39,6)) 
        		_xxrotin	:= 'IMPEDI01()'
        		/*
        		alert('PP1' +	nPedComp)
        		alert(nItemDoc)
        		alert(nCasDec)
        		alert(nPedComp)
        		alert(dDtaComp)
        		alert(cIdPedCo)
        		alert(dVigPedC) 
                */
        endif	
        
    	if (SUBSTR(cLine, 1,3))=='AE3'   
        		cLocFat	:=	(SUBSTR(cLine, 4,14))
        		cLocCob	:=	(SUBSTR(cLine, 18,14))
        		cLocEnt	:=	(SUBSTR(cLine, 32,14))
        		cTpTran	:=	(SUBSTR(cLine, 46,4))                             
        		cCnpjEnt:=	(SUBSTR(cLine, 50,14)) 
        		cFlxCarP:=  (SUBSTR(cLine, 64,1)) 
                /*
                alert('AE3')
   				alert(cLocFat)
        		alert(cLocCob)
        		alert(cLocEnt)
        		alert(cTpTran)
        		alert(cCnpjEnt)
        		alert(cFlxCarP)
                 */
                 _cdCliente:=""
	
				DbSelectArea("SA1")
				DbSetOrder(3)
				DbGoTop()

				if DbSeek(xFilial("SA1")+cLocFat)==.t.
					_cdCliente	:=	SA1->A1_COD   
					_cdLoja		:=	SA1->A1_LOJA
					_nmReduz	:=	SA1->A1_NREDUZ 
					_ufCli		:=	SA1->A1_EST
					_cVend		:=	SA1->A1_VEND
					_cTab		:=	SA1->A1_TABELA
					_cMun		:=	SA1->A1_MUN 
					_nomeC		 := SA1->A1_NOME 
				
	
				else
					alert("Ciente não encontrado! "+ cLocFat)//jb
					Return
				endif 
 
		
				DbSelectArea("SA3")
				DbSetorder(1)
				DbGoTop()
				If Dbseek(xfilial()+_cVend) == .T.
					_nomeV 		:= SA3->A3_NOME
					_a3email	:= SA3->A3_EMAIL
					
				else                     
					_nomeV := "Vendedor não encontrado"
				endif 
				
				
			if SA1->A1_MSBLQL == "1"
						//chamada de função para avisar que o cliente está bloqueado.   
						okEdISA1()
						Return
					
			endIf	   
	
				DbSelectArea("SC5")
				DbSetOrder(12)
				DbGoTop()
	
				If DbSeek(xfilial("SC5")+_cdCliente+_cdLoja+nPedComp+dDtaComp)==.T.
				  	alert("Pedido já existe cadastrado!!! "+nPedComp + " Nosso Pedido "+ c5_num)//jb
					_PedEdi:="S"
					Return
				endif	
       	endif
   	    	
   	    if (SUBSTR(cLine, 1,3))=='TE1'   
        		cTextL1	:=	(SUBSTR(cLine, 4,40))
        		cTextL2	:=	(SUBSTR(cLine, 44,40))
        		cTextL3	:=	(SUBSTR(cLine, 84,40))
        		cEspaco	:=	(SUBSTR(cLine, 124,5))      
        		if !empty(cTextL3)         
        		
        		_CPEDCLI:= encodeutf8(cTextL3)
        		
        			cMenNota :=_MNSCLI+ alltrim(_CPEDCLI)  
        		
        		endIf
                /*
                alert('AE3')
   				alert(cTextL1)
        		alert(cTextL2)
        		alert(cTextL3)
                */
				if _PedEdi =="N"
	               	okPedc5()
	            else
	            	Return
	            endif	   	
       	endif  

   	 
   	 	
		if (SUBSTR(cLine, 1,3))=='PP2'  
				cIdPrd	:=	(SUBSTR(cLine, 4,30))
				nQtdE	:=	(SUBSTR(cLine, 34,9))
				cDescIt	:=	(SUBSTR(cLine, 43,25))
				nVlItem	:=	(SUBSTR(cLine, 68,12))	
				cIdPed	:=	(SUBSTR(cLine, 86,4))
				dDtaEnt	:=	(SUBSTR(cLine, 90,6))
				cCdFisc	:=	(SUBSTR(cLine, 96,10))
				nAliqIPI:=	(SUBSTR(cLine, 106,4))
				nPDesc	:=	(SUBSTR(cLine, 110,4))
				cCdPgto	:=	(SUBSTR(cLine, 114,4))
				cEspac11:=	(SUBSTR(cLine, 118,11))
				
				if _PedEdi =="N"
					okPedc6() 
				else
					Return
				endif		
				/*
				alert(cItem)
				alert(nQtdE)
				alert(cDescIt)
				alert(nVlItem)
				alert(cIdPed)
				alert(dDtaEnt)
				alert(cCdFisc)
				alert(nAliqIPI)
				alert(nPDesc)
				alert(cCdPgto)
				alert(cEspac11)
			    */
			    _count:=_count+1
		endif	

    // LEITURA DA PROXIMA LINHA DO ARQUIVO
	    cBuffer := SPACE(lBuffer)                   // ALOCA BUFFER
   		lFilePos += lPos + 1                        // POSICIONA ARQUIVO APÓS O ULTIMO EOL ENCONTRADO
		lFilePos := FSEEK(hFile, lFilePos, 0)       // POSICIONA PONTEIRO DO ARQUIVO
	    lRead := FREAD(hFile, cBuffer, lBuffer)     // LE OS CARACTERES DO ARQUIVO
	    lPos := AT(cEOL, cBuffer) // PROCURA O PRIMEIRO FINAL DE LINHA

	END 
	
	IF _COUNT >= 1
		if _PedEdi =="N"
			okPedGrv()
			Processa({||okEdiEmail()})
		else
			Return
		endif	
	ENDIF	

RETURN    

Static function okPedc5()

	_lGravou  := .F.
	
	

	 

	//_auxPedido := "PEDIDO CLIENTE : " + aLLTRIM(C5_PEDCLI)

            
	aadd(aCabPV,{"C5_NUM"    	,cDoc			,Nil}) // Tipo de pedido                	
	aadd(aCabPV,{"C5_TIPO"    	,'N'	  		,Nil}) // Tipo de pedido  
	aadd(aCabPV,{"C5_CLIENTE" 	,_cdCliente 	,Nil}) // Codigo do cliente  				    
	aadd(aCabPV,{"C5_LOJACLI"	,_cdLoja		,Nil}) // Loja do cliente
	aadd(aCabPV,{"C5_CLIENT" 	,_cdCliente  	,Nil}) // Codigo do cliente  				    
	aadd(aCabPV,{"C5_LOJAENT"	,_cdLoja 		,Nil}) // Loja do cliente
	aadd(aCabPV,{"C5_XXROTIN"	,_xxrotin 		,Nil}) // Loja do cliente
	aadd(aCabPV,{"C5_EMISSAO"	,dDataBase  	,Nil})
	aadd(aCabPV,{"C5_XPEDEDI"	,nPedComp 	 	,Nil})
	aadd(aCabPV,{"C5_XDTAEDI"	,dDtaComp  		,Nil})
	aadd(aCabPV,{"C5_XTIPEDI"	,'E'	  		,Nil})   
	aadd(aCabPV,{"C5_OBSPED"	,cTextL1+' '+cTextL2		,Nil})      
	aadd(aCabPV,{"C5_MENNOTA"	,FwCutOff(alltrim(cMenNota)),Nil}) 
	
 
Return

//-------------------------------------------------------------------------------------------------
//  Itens
//-------------------------------------------------------------------------------------------------

Static function okPedc6()
   Local Int1 := 0
   Local Int2 := 0
   _cCdProd:=""
   
	DbSelectArea("SB1")
	dbSetOrder(14)
	DbGoTop()   
	_qtdPrd	:=	val(nQtdE) 
	If DbSeek(xfilial("SB1")+cIdPrd)==.T. 
		//_cCdProd:=SB1->B1_COD 
			if SB1->B1_XXMULTB>0
			 	int1 :=round(_qtdPrd/SB1->B1_XXMULTB,0)
			 	int2 :=(_qtdPrd/SB1->B1_XXMULTB)  
	   	  			If  Int1<>Int2   
                     
                      	nOKGrv()
                 		lMultB:=.T.  
                 		_cCdProd:=''   
                 		Return
                    else
                    _cCdProd:=SB1->B1_COD 
                    endIf
          		
            endIf
            If  SB1->b1_MSBLQL == "1"
            		nOKGrv() 
					lBlqProd:=.T.    
					_cCdProd:=''
			else
			_cCdProd:=SB1->B1_COD       
            EndIf           
            //
            if _cCdProd==''
                	lErroProd:=.T.
            endIf

	Else            
		DbSelectArea("SB1")
		dbSetOrder(1)
		DbGoTop()
		If DbSeek(xfilial("SB1")+cIdPrd)==.T. 
			_cCdProd:=SB1->B1_COD  
			cIdPrd	:=SB1->B1_CODCOM    
		
			
			
			
		ELSE
			nOKGrv()
		endif	
	Endif		

    
   	_it		:=	_it+1
 			if _it >= 98
 				if _it == 98
 					_cit :='AO'
 			    endif  
	 			    
	 			    _cit:= soma1(_cit,2)
	 			 //   alert(_cit)
	 			else                     
	 				_cit:= alltrim(StrZero(_it,2))
	 		endif
   
    IF _cCdProd ==""
    	nOKGrv()
    	Return
    endif	

	Aadd(aItemPV,{	{"C6_ITEM"   	,_cit					,Nil},;	//  numero do item   
					{"C6_NUM" 		,cDoc		   			,Nil},;	//  codigo do produto
					{"C6_PRODUTO"	,_cCdProd	   			,Nil},;	//  codigo do produto
					{"C6_CODCOM"	,cIdPrd					,Nil},;              
					{"C6_XXROTIN"	,_xxrotin 			 	,Nil},; // PEDIDO CLIENTE
					{"C6_ITEMPC"	,cIdPed				 	,Nil},; // PEDIDO CLIENTE
					{"C6_NUMPCOM"	,nPedComp		 		,Nil},; // PEDIDO CLIENTE
					{"C6_QTDVEN" 	,_qtdPrd		   		,Nil}})	//  quantidade vendida
  		
Return


static function okPedGrv()
	_lGravou:= .F.	
   	Begin Transaction
  		Processa( {||MSExecAuto({|x,y,z|Mata410(x,y,z)},aCabPv,aItemPV,3)},"Aguarde Gerando Pedido .... : " + cDoc )
   
		If lMsErroAuto
  		  	Mostraerro()
		  	DisarmTransaction()
		  	break
		Endif
   			_lGravou := .T.
	End Transaction

	If _lGravou
		ConfirmSX8()   
   	else
   	  	RollBackSx8() //Volta a numeração da Ordem de Liberação 
  	   	MSGBOX("PEDIDO COM PROBLEMAS VERIFIQUE !!!"  )//jb
	Endif    

Return

//Grava Tabela de inconsistencias de produto inexistente
Static Function nOKGrv()  
 
//	alert(xfilial()+" "+cLocFat+" "+nPedComp+" "+cIdPed+" "+cIdPrd+ " "+dDtaComp)
	DbSelectArea("ZZR")
	DbSetOrder(1)
	DbGoTop()
	If Dbseek(xfilial("  ")+cLocFat+nPedComp+cIdPed+cIdPrd+dDtaComp)==.T.
		Reclock("ZZR",.F.)	
//		alert("Encontrou ")
	ELSE                  
//		alert("Não Encontrou")
		Reclock("ZZR",.T.)
	ENDIF
//ZZR_FILIAL+ZZR_CLIENT+ZZR_LOJA+ZZR_PROD+ZZR_PEDCLI+ZZR_ITEDI+ZZR_DATA                                                                                                                                                                                    
	_itN := _itN + 1
	ZZR_CLIENT  := _cdCliente
	ZZR_LOJA    := _cdLoja
	ZZR_PROD    :=	cIdPrd
	ZZR_NMCLI   :=  _nmReduz 
	ZZR_OBS     :=	"inconsistencia nosso pedido " +cDoc +" "+DTOS(DDATABASE) "
	ZZR_PEDCLI  :=	nPedComp
	ZZR_DATA    :=	DDATABASE 
	ZZR_DTAPED	:=	dDtaComp
	ZZR_ITEDI   :=	cIdPed
	ZZR_FILE    :=	alltrim(_aFiles[i,1]) 
	ZZR_UFCLI	:=	_ufCli	   
	ZZR_VEND	:=	_cVend
	ZZR_TABPRC	:=	_cTab
	ZZR_MUNCLI	:=	_cMun  
	ZZR_CNPJ	:=	cLocFat 
	Msunlock()  
	
Return	          

Static Function okEdiEmail() 
 
//LOCAL bCondic := { | | A1_COD >= “000001” .AND. A1_COD <= “001000” }  

	if SM0->M0_CODIGO == '01'                                    
 		_emite	:=	"GAUSS IND. E COM. DE AUTO PECAS LTDA.   "
	elseIF SM0->M0_CODIGO == '02'                
	 	_emite	:=	"CDG COMPONENTES AUTOMOTIVOS LTDA        "
	elseIF SM0->M0_CODIGO == '04'                
	 	_emite	:=	"NSG COMPONENTES AUTOMOTIVOS LTDA        "
	endif 
 
	lEnvia    := .T.
	cAccount  := AllTrim(GetMV("MV_NFECDGE"))
	cPass     := AllTrim(GetMV("MV_NFECDGP"))
	cServer   := AllTrim(getMV("MV_NFECDGS"))
	cUser     := AllTrim(getMV("MV_NFECDGU"))  
	cComercial:= AllTrim(getMV("MV_CMEMAIL"))                                      
	cMensagem   := ''
	aLista      := {}
	_CRLF       := chr(13)+chr(10)
	cMSG        := space(200)
	aFiles      := {}
   	nI          := 1  
   	cPara     := trim(_a3email)  
 //  	cPara     := "loliveira-pauloh@hotmail.com"
   //	cPara     := "nilson.goncalves@gauss.ind.br"
	_cCC      :=   "paulo.enigma@gauss.ind.br"+";"+"nilson.goncalves@gauss.ind.br"+";"+"comercialcdg@cdglogistica.com.br"
	Private aDados    := {}
	_dDtaComp1	:= substr(dDtaComp,5,2)+"/"+substr(dDtaComp,3,2)+"/"+substr(dDtaComp,1,2) 
	_itok	:=	alltrim(str(_it))
	_itNok	:=	alltrim(str(_itN))   
	//APLICAR Valor final
	_VlrFInal := 0   
	 TRT01     := GetNextAlias() 

 
BeginSql Alias TRT01
		SELECT SUM(C6_VALOR)Valor_fim 
	FROM %table:SC6% C6
	WHERE C6.D_E_L_E_T_=''
	AND C6_NUM = %exp:SC5->C5_NUM% 
 EndSql    
 
 While !(TRT01)->(EOF()) 		
  
_VlrFInal:= alltrim(str((TRT01)->Valor_fim))
		                                  
	(TRT01)->(DbSkip())
EndDo 


                                           
	
  
	cMsg += '<html>' + _CRLF
   	cMsg += '<title>EMITENTE '+ _EMITE + '</title>'+_CRLF 
 	cMsg += '</head>' + _CRLF  
  	cMsg += '<font size="2" face="Square721 Blk" COLOR="#0070C0"> PEDIDOS EDI, '+nPedComp + ' -  '+_dDtaComp1 +" - "+ substr(_nmReduz,1,25)+ ' '+ _CRLF 
    cMsg += '<hr color=navy size=5>'+ _CRLF  
 	cMsg += '<Table bgcolor="#F0F8FF"'//style="border: 1px #003366 solid; border="0" cellpadding ="1" cellspacing="2" width="553" height="600">'  
	cMsg += '<font size="2" face="Arial" COLOR="#C00000" > R E S U M O</font><BR>' + _CRLF                   	
 	cMsg += '<font size="2" face="Arial" COLOR="#C00000">EDI				: ' +nPedComp + ' EMISSAO  :' + DTOC(SC5->C5_EMISSAO) + '</font><BR>' + _CRLF                   
	cMsg += '<font size="2" face="Arial" >REPRESENTANTE	: ' +_nomeV + '</font><BR>' + _CRLF   
	cMsg += '<font size="2" face="Arial" >CLIENTE 			: '+substr(_nomeC,1,25) + '</font><BR>' + _CRLF   
  	cMsg += '<font size="2" face="Arial">CNPJ 			: '+SA1->A1_CGC + '</font><BR>' + _CRLF   
  	cMsg += '<font size="2" face="Arial">NOSSO PEDIDO 	: '+SC5->C5_NUM + ' | EMITIDO | '+ DTOC(SC5->C5_EMISSAO)  +'</font><BR>'  + _CRLF   
    cMsg += '<font size="2" face="Arial">MENSAGEM 		: '+ TRIM(cTextL1) + '</font><BR>' + _CRLF 
   	cMsg += '<font size="2" face="Arial">COND.PAGTO.	: '+SA1->A1_COND + "  " +'</font><BR>' + _CRLF   
  	cMsg += '<font size="2" face="Arial">TABELA			: '+SA1->A1_TABELA +'</font><BR>' + _CRLF 
   	cMsg += '<hr color=navy>'+ _CRLF  
  	cMsg += '<font size="2" face="Arial" COLOR="#C00000" > P R O D U T O S</font><BR>' + _CRLF 
  	cMsg += '<font size="2" face="Arial">Qtd Itens importados		: '+_itok +'</font><BR>' + _CRLF   
  	cMsg += '<font size="2" face="Arial">Valor Total do Pedido		: '+_VlrFInal +'</font><BR>' + _CRLF
  	cMsg += '<font size="2" face="Arial">Qtd Itens inconsistentes	: '+_itNok +'</font><BR>' + _CRLF   
  	//---- 
 
  ZZR->(DbCloseArea())
if   lMultB == .T. .or. lBlqProd==.T. .or. lErroProd ==.T.
DbSelectArea("ZZR")
set filter to ZZR->ZZR_PEDCLI==nPedComp
dbSetOrder(1)
dbGoTop()
//dbSetFilter(,bCondic)  
//dbGoBotton()
	if lMultB == .T.
   		cMsg += '<font size="2" face="Arial">Produtos com inconsistência de múltiplos	: '+_itNok +'</font><BR>' + _CRLF 
   	 elseIf lBlqProd==.T.
   		cMsg += '<font size="2" face="Arial">Produtos Bloqueados	: '+_itNok +'</font><BR>' + _CRLF  
     else 
     	cMsg += '<font size="2" face="Arial">Produtos com códigos errados	: '+_itNok +'</font><BR>' + _CRLF  

   	endIf
   	  
	while    !ZZR->(EOF()) 
			cMsg += '<font size="2" face="Arial">Código	:'+ZZR->ZZR_PROD +' </font><BR>' + _CRLF   
			dbSkip()
    endDo
endIf
//----
  		
	cMsg += '<font size="2" face="Arial" COLOR="#C00000" > CONSULTAR ANALISE : 346 - Representantes_Pedidos_EDI</font><BR>' + _CRLF                   	
	cMsg += '</head>' + _CRLF
	cMsg += '<hr color=navy>'+ _CRLF  
	cMsg += '<font size="2" face="Arial"> Não precisa responder este email, é somente um aviso sobre recebimento do seu pedido de compra.' + _CRLF				
	cMsg += '<font size="2" face="Arial"> Muito obrigado! '+ _CRLF 
	cMsg += '</head>' + _CRLF 
	cMsg += '<font size="2" face="Arial"> Tel. Comercial 41 35526075 '+ _CRLF 
	cMsg += '</head>' + _CRLF


  
	Aadd( aLista, cPara )
	CONNECT SMTP SERVER cServer ACCOUNT cAccount PASSWORD cPass   RESULT lConectou
	IF lConectou
	    lRet := Mailauth(cUser,cPass)
        Reclock("SC5",.f.)
		C5_XXEMAIL := "S"
		Msunlock()

	    If lRet  

	    	SEND MAIL FROM cAccount ;
		    TO aLista[1] ;
		    CC _cCC  ;
		    SUBJECT   'E-MAIL AUTOMATICO - AVISO PEDIDO EDI ' + SC5->C5_XPEDEDI + ' |  ' +substr(_nomeC,1,25) ; //+ '| Emitente ' + _emite;
		    BODY cMsg RESULT lOK
	 //	    ALERT("Enviado sucesso")
		ElseIf !lRet
	       //	alert('Não authenticado ')  //
		Endif
	    DISCONNECT SMTP SERVER
	ELSEIF (! lConectou)
	   GET MAIL ERROR cERRO
	  // alert('Erro de conexão '+cERRO)//
	ENDIF
 	
Return             

//Função para avisa que CLiente X está bloqueado
Static Function okEdISA1()

	if SM0->M0_CODIGO == '01'                                    
 		_emite	:=	"GAUSS IND. E COM. DE AUTO PECAS LTDA.   "
	else              
	 	_emite	:=	"CDG COMPONENTES AUTOMOTIVOS LTDA        "
	endif 
 
	lEnvia    := .T.
	cAccount  := AllTrim(GetMV("MV_NFECDGE"))
	cPass     := AllTrim(GetMV("MV_NFECDGP"))
	cServer   := AllTrim(getMV("MV_NFECDGS"))
	cUser     := AllTrim(getMV("MV_NFECDGU"))  
	cComercial:= AllTrim(getMV("MV_CMEMAIL"))                                      
	cMensagem   := ''
	aLista      := {}
	_CRLF       := chr(13)+chr(10)
	cMSG        := space(200)
	aFiles      := {}
   	nI          := 1  
   	cPara     := trim(_a3email)  
 //  	cPara     := "loliveira-pauloh@hotmail.com"
  //	cPara     := "nilson.goncalves@gauss.ind.br"
	_cCC      :=  "paulo.enigma@gauss.ind.br"+";"+"nilson.goncalves@gauss.ind.br"+";"+"comercialcdg@cdglogistica.com.br"
	Private aDados    := {}
	_dDtaComp1	:= substr(dDtaComp,5,2)+"/"+substr(dDtaComp,3,2)+"/"+substr(dDtaComp,1,2) 
	_itok	:=	alltrim(str(_it))
	_itNok	:=	alltrim(str(_itN))    
	cMsg += '<html>' + _CRLF
   	cMsg += '<title>EMITENTE '+ _EMITE + '</title>'+_CRLF 
 	cMsg += '</head>' + _CRLF  
  	cMsg += '<font size="2" face="Square721 Blk" COLOR="#0070C0"> PEDIDOS EDI, '+nPedComp + ' -  '+_dDtaComp1 +" - "+ substr(_nmReduz,1,25)+ ' '+ _CRLF 
    cMsg += '<hr color=navy size=5>'+ _CRLF  
 	cMsg += '<Table bgcolor="#F0F8FF"'//style="border: 1px #003366 solid; border="0" cellpadding ="1" cellspacing="2" width="553" height="600">'  
	cMsg += '<font size="2" face="Arial" COLOR="#C00000" > R E S U M O</font><BR>' + _CRLF
	cMsg += '<font size="2" face="Arial" >REPRESENTANTE	: ' +_nomeV + '</font><BR>' + _CRLF   
	cMsg += '<font size="2" face="Arial" > CLIENTE 			: '+substr(_nomeC,1,25) + ' ESTA BLOQUEADO </font><BR>' + _CRLF   
  	cMsg += '<font size="2" face="Arial">CNPJ 			: '+SA1->A1_CGC + '</font><BR>' + _CRLF  
	cMsg += '<font size="2" face="Arial"> Por gentileza entrar em contato com o comercial para verificar situação.' + _CRLF		
	cMsg += '<font size="2" face="Arial">Após o Desbloqueio do Cliente enviar o pedido novamente.' + _CRLF		
  
   	cMsg += '<hr color=navy>'+ _CRLF                    	
	cMsg += '</head>' + _CRLF
	cMsg += '<hr color=navy>'+ _CRLF  
	cMsg += '<font size="2" face="Arial"> Não precisa responder este email, é somente um aviso sobre recebimento do seu pedido de compra.' + _CRLF				
	cMsg += '<font size="2" face="Arial"> Muito obrigado! '+ _CRLF 
	cMsg += '</head>' + _CRLF 
	cMsg += '<font size="2" face="Arial"> Tel. Comercial 41 35526075 '+ _CRLF 
	cMsg += '</head>' + _CRLF


  
	Aadd( aLista, cPara )
	CONNECT SMTP SERVER cServer ACCOUNT cAccount PASSWORD cPass   RESULT lConectou
	IF lConectou
	    lRet := Mailauth(cUser,cPass)
     //   Reclock("SC5",.f.)
	//	C5_XXEMAIL := "S"
	//	Msunlock()

	    If lRet  

	    	SEND MAIL FROM cAccount ;
		    TO aLista[1] ;
		    CC _cCC  ;
		    SUBJECT   'E-MAIL AUTOMATICO - AVISO DE Inconsistências do Cliente: ' +substr(_nomeC,1,25) ; //+ '| Emitente ' + _emite;
		    BODY cMsg RESULT lOK
	 //	    ALERT("Enviado sucesso")
		ElseIf !lRet
	      //	alert('Não authenticado ')//
		Endif
	    DISCONNECT SMTP SERVER
	ELSEIF (! lConectou)
	   GET MAIL ERROR cERRO
	  // alert('Erro de conexão '+cERRO) //
	ENDIF
 	


Return

Static Function okold()

	copy File &_carqTxt to &_carqOld
	
	IF FERASE(_carqTxt) == -1
     //	MsgStop('Falha na deleção do Arquivo ( FError'+str(ferror(),4)+             ')')   //
	Else
//	    MsgStop('Arquivo deletado com sucesso.')
	ENDIF

Return#   R E A D M E  
 
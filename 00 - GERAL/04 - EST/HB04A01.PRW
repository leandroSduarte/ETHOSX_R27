#Include 'TOTVS.ch'
#include 'parmtype.ch'
#Include 'protheus.ch'
#Include 'FileIO.ch'

#define CMD_OPENWORKBOOK			1
#define CMD_CLOSEWORKBOOK		   	2
#define CMD_ACTIVEWORKSHEET  		3
#define CMD_READCELL				4

//-------------------------------------------------------------------
/*/{Protheus.doc} HB04A01
description Leitura e importaçao de dados via planilha para criacao de
            movimentações internas.

            Uso: Contábil Habibs
            pastel 1
@author  Leonardo Espinosa
@since   16/01/2019
@version 1.0
/*/
//-------------------------------------------------------------------
User Function HB04A01( )
Local aArea		:= GetArea( )

Local aParams	:= { }
Local aHeader   := { }
Local aDados    := { }

Private _aRet	:= { }
Private _cFile	:= SPACE(500)
Private _CTITULO:= "Importação de ganhos"

aAdd( aParams  ,{;
	6,;                                                 //[1]MsGet
	"Arquivo",;                                         //[2]Descrição
	_cFile,;                                            //[3]String contendo o inicializador do campo
	"",;                                                //[4]String contendo a Picture do campo
	,;                                                  //[5]String contendo a validação
	"",;							                    //[6]String contendo a validação When
	90,;                                                //[7]Tamanho do MsGet
	.T.,;                                               //[8]Flag .T./.F. Parâmetro Obrigatório ?
	"Arquivos .CSV |*.CSV",;                   //[9]Texto contendo os tipos de arquivo Ex.: "Arquivos .CSV |*.CSV"
		"C:\",;                                         //[10]Diretório inicial do cGetFile
		GETF_LOCALHARD+GETF_LOCALFLOPPY+GETF_NETWORKDRIVE;  //[11]Parâmetros do cGetFile
	})


//-- Abre a caixa de parametros
If ParamBox(aParams,"Parametros",@_aRet,/*bOk*/ ,/*aButtons*/ ,.T. ,0,0,/*oDlgWizard*/ ,/*cLoad*/ , .F., .F.)

    _cFile := _aRet[01]
    
    If !File(_cFile)
        MsgInfo('A planilha não foi localizada.' + CRLF + CRLF;
        + 'A rotina não será processada.'+ CRLF + CRLF;
        + 'Favor verificar.',;
        + _cTitulo)
    Else
        _cFile := Alltrim(_cFile)

		If '.XLS' $ _cFile
        	Processa({|| OpenExcel(_cFile, @aHeader, @aDados), "Processando dados da planilha, não feche o arquivo..."})
		Else 
			Processa({|| OpenCSV(_cFile, @aHeader, @aDados), "Processando dados do CSV..."})
			

		EndIf 
        //Processa os dados
        If Len(aDados) == 0
            MsgAlert("Atenção!"+CRLF;
                    +"Nenhuma linha foi lida dentro de arquivo. Confira se o layout esta correto.")

        Else
            Processa({|| HB04A01A(@aDados, aHeader), "Processando dados..."}) 
            

        EndIf 
    EndIf 

Else
	MsgAlert("Execução cancelada pelo usuário!")
	Return
EndIf

RestArea(aArea)	


Return


//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Abre e efetua a leitura do arquivo excel					 ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
Static Function OpenExcel(_cFile, aHeader, aDados)

Local _cBuffer    := ''          //variável de apoio a abertura da planilha em Excel
Local _cWorksheet := 'Planilha1' 	//Informe a Worksheet (aba) do Excel que será lida pelo programa
Local _cColunas   := 'A|B|C|D|E|F|G|H|I|J|K|L|M|N|O|P|Q|R|S|T|U|V|W|X|Y|Z|AA|AB|AC|AD|AE|AF|AG|AH|AI|AJ|AK|AL|AM|AN|AO|AP|AQ|AR|AS|AT|AU|AV|AW|AX|AY|AZ|BA|BB|BC|BD|BE|BF|BG|BH|BI|BJ|BK|BL|BM|BN|BO|BP|BQ|BR|BS|BT|BU|BV|BW|BX|BY|BZ'
Local _aColunas   := STRTOKARR(_cColunas,"|")
Local _cLastColun := 'E'
Local _aReadColun := {}
Local _aCabec     := {}
Local _aRefCabec  := {}		
Local _aItens     := {}
Local _cItens     := ''
Local _aAuxItens  := {}
Local _nLinha     := 0
Local _nColuna    := 0
Local _nLinCabec  := 1       //Informe a linha do cabeçalho
Local _nLinFim    := 500000  //Informe a linha final do arquivo a ser lido
Local _nHdl       := 0
Local _cPath      := '\DLL\'//Informe o caminho do servidor onde está a Dll de leitura do arquivo em Excel
Local _cDll       := 'readexcel.dll'
Local _cTemp      := 'C:\TEMP\'
Local _nA         := 0
Local _nB         := 05
Local _nRead      := 0


    _cLastColun := 'O'
    _cWorksheet := 'Ajuste'
    _nLinCabec	:= 1

	//+------------------------------------------------------------------------------+
	//| Tratamento para copiar a DLL do Servidor e salvar local na máquina do usuário|
	//+------------------------------------------------------------------------------+
	If ExistDir(_cPath)
		If Directory(_cPath+_cDll)[1][1] == Upper(_cDll)
			If !ExistDir(_cTemp)
				If MakeDir(_cTemp) == 0
					CpyS2T(_cPath+_cDll,_cTemp,.T.) // COPIA ARQUIVO PARA MAQUINA DO USUÁRIO	
				Endif	
			Else
				FERASE(_cTemp+_cDll)
				CpyS2T(_cPath+_cDll,_cTemp,.T.) // COPIA ARQUIVO PARA MAQUINA DO USUÁRIO
			Endif
		Endif
	Endif

	//+------------------------------------------------------------------------------+
	//| Efetua a abertura da DLL de leitura de planilhas em Excel                    |
	//+------------------------------------------------------------------------------+
	_nHdl       := ExecInDLLOpen(_cTemp+_cDll)

	If (_nHdl >= 0)

		ProcRegua(2)
		//+------------------------------------------------------------------------------+
		//| Carrega o Excel e abre a planilha                                            |
		//+------------------------------------------------------------------------------+
		_cBuffer := _cFile + Space(512)
		nBytes   := ExeDLLRun2(_nHdl, CMD_OPENWORKBOOK, @_cBuffer)

		If (nBytes < 0)
			//+------------------------------------------------------------------------------+
			//| Erro critico na abertura do arquivo sem mensagem de erro                     |
			//+------------------------------------------------------------------------------+
			MsgStop('Não foi possível abrir o arquivo : ' + cFile)
		ElseIf (nBytes > 0)
			//+------------------------------------------------------------------------------+
			//| Erro critico na abertura do arquivo com mensagem de erro                     |
			//+------------------------------------------------------------------------------+
			MsgStop(Subs(_cBuffer, 1, nBytes))
			ExeDLLRun2(_nHdl, CMD_CLOSEWORKBOOK, @_cBuffer)
			ExecInDLLClose(_nHdl)
		EndIf

		//+------------------------------------------------------------------------------+
		//| Leitura da WorkSheet da planilha em Excel                                    |
		//+------------------------------------------------------------------------------+
		_cBuffer := _cWorksheet
		ExeDLLRun2(_nHdl, CMD_ACTIVEWORKSHEET, @_cBuffer)

		///+------------------------------------------------------------------------------+
		//| Montagem do Cabeçalho                                                        |
		//+------------------------------------------------------------------------------+
		For _nColuna := 1 to Len(_aColunas)
			IncProc("Não feche a planilha!"+CHR(10)+CHR(13)+" Lendo Aba: "+_cWorksheet+", Celula: "+_aColunas[_nColuna]+cValtoChar(_nLinCabec))
			_cDados := ReadCell(_nHdl,_aColunas[_nColuna],_nLinCabec)		
			If (_cLastColun == _aColunas[_nColuna])
				Exit
			Else
				Aadd(_aReadColun,_aColunas[_nColuna])
				Aadd(_aCabec,_cDados)
			EndIf
		Next _nColuna
		
		
		//+------------------------------------------------------------------------------+
		//| Validação das colunas do Cabeçalho                                           |
		//+------------------------------------------------------------------------------+

        Aadd(_aRefCabec,{'Filial'							,' ' 	})
        Aadd(_aRefCabec,{'TP Movimento'						,' ' 	})
        Aadd(_aRefCabec,{'Produto'					        ,' '	})
        Aadd(_aRefCabec,{'Unidade'				    		,' '	})
        Aadd(_aRefCabec,{'Quantidade'   					,' '	})
        Aadd(_aRefCabec,{'Tipo RE/DE'	                    ,' '	})
        Aadd(_aRefCabec,{'C Contabil'						,' '	})
        Aadd(_aRefCabec,{'Armazem'							,' '	})
        Aadd(_aRefCabec,{'Documento'    					,' '	})
        Aadd(_aRefCabec,{'DT Emissao'       				,' '	})
        Aadd(_aRefCabec,{'Grupo'        					,' '	})
        Aadd(_aRefCabec,{'Custo'        					,' '	})
        Aadd(_aRefCabec,{'Tipo'        					    ,' '	})
        Aadd(_aRefCabec,{'Descr. Prod'  				    ,' '	})
		
		
		For _nA := 1 to Len(_aRefCabec)
			For _nB := 1 to Len(_aCabec)
				If UPPER(Alltrim(_aRefCabec[_nA][1])) == UPPER(Alltrim(_aCabec[_nB]))
					_aRefCabec[_nA][2] := 'X'
					
					_nA++
				Endif
			Next _nB
		Next _nA

		For _nA := 1 to Len(_aRefCabec)
			If !(_aRefCabec[_nA][2] == 'X')
				MsgStop('Existem divergências entre o cabeçalho da planilha e o cabeçalho configurado da rotina denominada '+_cTitulo+'.' + CRLF + CRLF;
				+ 'A rotina não será processada.'+ CRLF + CRLF;
				+ 'Favor verificar.',;
				+ _cTitulo)	
				_aCabec := {}
				Exit
			Endif
		Next _nA
		
		//+------------------------------------------------------------------------------+
		//| Montagem dos dados - itens                                                   |
		//+------------------------------------------------------------------------------+
		If Len(_aCabec) > 0
			For _nA := (_nLinCabec+1) To _nLinFim
				_aAuxItens := {}
				_cItens := ''
				For _nB := 1 To Len(_aReadColun)
					IncProc("Não feche a planilha! Lendo Aba: "+_cWorksheet+", Celula: "+_aReadColun[_nB]+cValToChar(_nA))
					_cDados := ReadCell(_nHdl,_aReadColun[_nB],_nA)					
					_cItens += _cDados+" |"
				Next _nB
				_aAuxItens := STRTOKARR(_cItens,"|")
				If !Empty(Alltrim(_aAuxItens[1]))			
					Aadd(_aItens,_aAuxItens)
				Else
					_nRead := _nA-1
					Exit
				EndIf
			Next _nA
		EndIf
		If _nRead == 0
			_nRead := _nA
		Endif  

		aHeader := Aclone(_aCabec)
		aDados  := Aclone(_aItens)

		//+------------------------------------------------------------------------------+
		//| Fecha o arquivo e remove o excel da memória                                  |
		//+------------------------------------------------------------------------------+
		_cWorksheet := Space(512)
		ExeDLLRun2(_nHdl, CMD_CLOSEWORKBOOK, @_cWorksheet)
		ExecInDLLClose(_nHdl)
		_cFile     := SUBSTR(_cFile, RAT( "\"   , _cFile ),Len(_cFile))
		_cFile     := StrTran(_cFile,"\","")		
		If Len(aDados) > 0
			MsgInfo('Realizada a leitura de '+Alltrim(STR((_nRead)))+' linhas referente a planilha:' + CRLF + CRLF;
			+ 'Planilha: ' +_cFile + CRLF + CRLF;
			+ '',;
			+ _cTitulo)
   
        EndIf
	Else
		MsgStop('Não foi possivel carregar a DLL para a leitura da planilha.' + CRLF + CRLF;
		+ 'DLL: '+_cDll+ CRLF + CRLF;
		+ 'Favor verificar.',;
		+ _cTitulo)							 		
	EndIf
		
Return

/*/
+------------------------------------------------------------------------------+
|                           FICHA TÉCNICA DO PROGRAMA                          |
+------------------------------------------------------------------------------+
|Função            | ReadCell                                                  |
+------------------------------------------------------------------------------+
|Descrição         | Função utilizada para realizar a leitura do conteúdo da   |
|                  | célula da planilha de Excel                               |
+------------------------------------------------------------------------------+
|Autor             | Leonardo Espinosa      -     RVG                          |
+------------------------------------------------------------------------------+
|Data de Criação   | 26/01/2018                                                |
+------------------------------------------------------------------------------+
|Histórico de      |                                                           |
|Atualizações      |                                                           |
|                  |                                                           |
+------------------------------------------------------------------------------+
/*/
Static Function ReadCell(_nArq,_nCol,_nLinha)

	Local _cConteudo := ''
	Local _cBufferPl := ''
	Local _cCelula	 := ''

	//+------------------------------------------------------------------------------+
	//| Efetua a leitura da célula                                                   |
	//+------------------------------------------------------------------------------+
	_cCelula    := _nCol+Alltrim(str(_nLinha))
	_cBufferPl  := _cCelula + Space(1024)
	_nBytesPl   := ExeDLLRun2(_nArq, CMD_READCELL, @_cBufferPl)
	_cConteudo  := Subs(_cBufferPl, 1, _nBytesPl)
	_cConteudo  := Alltrim(_cConteudo)

Return (_cConteudo)

//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
//³ Processa os dadoso dos produtos 							 ³
//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
Static Function HB04A01A(aDados, aHeader)
Local   aAux    := { }
Local   aErro   := { }
Local   aLog    := { }
Local	aArea	:= { }
Local 	aCompany:= { }

Local   cMsgErr := ""
Local   cObs    := ""
Local	cFilBkp	:= ""
Local	cAlias	:= ""
Local 	cQry	:= ""

Local   nX      := 0
Local   nY      := 0
Local   nErro   := 0
Local 	nHndl	:= 0
Local 	nHndLog	:= 0

Local   nPosFil := aScan(aHeader,{|x| Capital(Alltrim(x)) == "Filial"       })
Local   nPosTM  := aScan(aHeader,{|x| Capital(Alltrim(x)) == "Tp Movimento" })
Local   nPosProd:= aScan(aHeader,{|x| Capital(Alltrim(x)) == "Produto"      })
Local   nPosUM  := aScan(aHeader,{|x| Capital(Alltrim(x)) == "Unidade"      })
Local   nPosQtd := aScan(aHeader,{|x| Capital(Alltrim(x)) == "Quantidade"   })
Local   nPosCF  := aScan(aHeader,{|x| Alltrim(x)          == "Tipo RE/DE"   })
Local   nPosCta := aScan(aHeader,{|x| Capital(Alltrim(x)) == "C Contabil"   })
Local   nPosAz  := aScan(aHeader,{|x| Capital(Alltrim(x)) == "Armazem"      })
Local   nPosDoc := aScan(aHeader,{|x| Capital(Alltrim(x)) == "Documento"    })
Local   nPosDt  := aScan(aHeader,{|x| Capital(Alltrim(x)) == "Dt Emissao"   })
Local   nPosGrp := aScan(aHeader,{|x| Capital(Alltrim(x)) == "Grupo"        })
Local   nPosCus := aScan(aHeader,{|x| Capital(Alltrim(x)) == "Custo"        })
Local   nPosTip := aScan(aHeader,{|x| Capital(Alltrim(x)) == "Tipo"         })

Private lMSErroAuto := .F.
Private lMsHelpAuto	:= .T.
Private lAutoErrNoFile := .T.

ProcRegua(Len(aDados)+1)

cMsgErr := ""
cObs    := "HB04A01-"+DTOS(dDataBase)+"-"+Time()

aCompany	:= ADMAbreSM0()

Begin Transaction
	For nX := 1 to Len(aDados)
		aArea	:= GetArea( )

		IncProc("Processando linha "+cValToChar(nX) + " Arquivo "+_cFile+ " Processado: " + cValToChar( ROUND( (nX/Len(aDados))*100,2) ) + "%") 
		
		aDados[nX,nPosFil] 	:= Alltrim(aDados[nX,nPosFil])
		aDados[nX,nPosQtd]	:= VAL(StrTran(aDados[nX,nPosQtd],",","."))
		aDados[nX,nPosCus]	:= VAL(StrTran(aDados[nX,nPosCus],",","."))

		cAlias	:= GetNextAlias( )
		If Select(cAlias) > 0
			(cAlias)->(DBCloseArea( ))
		EndIf 

		Begin Sequence 
			BeginSQL Alias cAlias
				SELECT
					D3_NUMSEQ
				FROM
					%Table:SD3% SD3 
				WHERE	
					D3_FILIAL 	= %Exp:aDados[nX,nPosFil]% 		AND 
					D3_DOC 		= %Exp:aDados[nX,nPosDoc]%		AND
					D3_COD 		= %Exp:aDados[nX,nPosProd]%		AND
					D3_QUANT	= %Exp:aDados[nX,nPosQtd]%		AND
					D3_TM		= %Exp:aDados[nX,nPosTM]%		AND 
					D3_EMISSAO 	= %Exp:dtos(CtoD(aDados[nX,nPosDt]))%	AND 
					SD3.%NotDel% 
			EndSQL
		Recover 
			cQry	:= GetLastQuery()[2]

		End Sequence

		DBSelectArea(cAlias)
		If (cAlias)->(!EOF( ))
			nErro++
			cMsgErr += _cFile+" Linha: "+cValToChar(nX)+ "| Documento "+ aDados[nX,nPosDoc] + " | Produto "+ aDados[nX,nPosProd] + " | Qtd.:"+ cValtoChar(aDados[nX,nPosQtd])+" | TM: "+aDados[nX,nPosTM]+" | Emissao: "+(aDados[nX,nPosDt])+CRLF+CRLF

			cMsgErr	+= "Linha ja inclusa anteriormente / DOCSEQ "+(cAlias)->D3_NUMSEQ

			cMsgErr += CRLF+CRLF+Replicate("=",50)+CRLF+CRLF

			Loop 
		EndIf 

		(cAlias)->(DBCloseArea( ))

		aDados[nX,nPosFil]	:= PADL(aDados[nX,nPosFil], TamSX3("D3_FILIAL")[01], "0")

		If cEmpAnt+cFilAnt <> cEmpAnt+aDados[nX,nPosFil]

			If aScan(aCompany, {|x| x[01] == cEmpAnt .AND. x[02] == aDados[nX,nPosFil] }) == 0
				nErro++
				cMsgErr += _cFile+" Linha: "+cValToChar(nX)+ "| Documento "+ aDados[nX,nPosDoc] + " | Produto "+ aDados[nX,nPosProd] + " | Qtd.:"+ cValtoChar(aDados[nX,nPosQtd])+" | TM: "+aDados[nX,nPosTM]+" | Emissao: "+CtoD(aDados[nX,nPosDt])+CRLF+CRLF

				cMsgErr	+= "Codigo de filial nao encontrada dentro da empresa atual. Empresa: "+cEmpAnt+" | Filial: "+aDados[nX,nPosFil]

				cMsgErr += CRLF+CRLF+Replicate("=",50)+CRLF+CRLF		
					
				Loop
			EndIf 

			cFilBkp	:= cFilAnt

			cFilAnt := aDados[nX,nPosFil]
			DBCloseAll( )

			OpenFile(cEmpAnt+aDados[nX,nPosFil])

		Endif

		aAux    := { }
		aSize(aAux, Len(aAux))

		DBSelectArea("SB2")
		DBSetOrder(1)
		If SB2->(!dbSeek(aDados[nX,nPosFil]+PADR(AllTrim(aDados[nX,nPosProd]),TamSX3("B1_COD")[01])+aDados[nX,nPosAz]))
			// OpenFile(cEmpAnt+cFilAnt)
			CriaSB2(aDados[nX,nPosProd],aDados[nX,nPosAz])

		EndIf

		SB2->(DBCloseArea( ))

		aDados[nX,nPosTM]	:= PADL(aDados[nX,nPosTM], TamSX3("D3_TM")[01], "0")
		aDados[nX,nPosAz]	:= PADL(aDados[nX,nPosAz], TamSX3("D3_LOCAL")[01], "0")

		aAux := {   {"D3_FILIAL"    ,aDados[nX,nPosFil]                 , NIL   },;
					{"D3_DOC"       ,aDados[nX,nPosDoc]                 , NIL   },;
					{"D3_TM"        ,aDados[nX,nPosTM]                  , NIL   },;
					{"D3_EMISSAO"   ,CtoD(aDados[nX,nPosDt])            , NIL   },;
					{"D3_COD"       ,aDados[nX,nPosProd]                , NIL   },;
					{"D3_UM"        ,aDados[nX,nPosUM]                  , NIL   },;
					{"D3_QUANT"     ,aDados[nX,nPosQtd]                 , NIL   },;
					{"D3_TIPO"      ,aDados[nX,nPosTip]                 , NIL   },;
					{"D3_CONTA"     ,aDados[nX,nPosCta]                 , NIL   },; 
					{"D3_LOCAL"     ,aDados[nX,nPosAz]                  , NIL   },;
					{"D3_GRUPO"     ,aDados[nX,nPosGrp]                 , NIL   },;
					{"D3_OBSERVA"   ,cObs                               , NIL   } }

		If aDados[nX,nPosCus] > 0
			aAdd(aAux,  {"D3_CUSTO1"    ,aDados[nX,nPosCus]                 , NIL   } ) 
		EndIf 

		lMSErroAuto	:= .F.
		MSExecAuto({|x,y| mata240(x,y)},aAux,3)

		If lMSErroAuto 
			cMsgErr += _cFile+" Linha: "+cValToChar(nX)+ "| Documento "+ aDados[nX,nPosDoc] + " | Produto "+ aDados[nX,nPosProd] + " | Qtd.:"+CvaltoChar(aDados[nX,nPosQtd])+CRLF+CRLF

			aErro   := GetAutoGRLog()
			For nY := 1 to Len(aErro)
				If 'AJUDA' $ aErro[nY] .OR. 'Invalido' $ aErro[nY]

					cMsgErr += aErro[nY]+CRLF
				
				EndIf 

			Next nY

			nErro++
			cMsgErr += CRLF+CRLF+Replicate("=",50)+CRLF+CRLF
		Endif

		RestArea(aArea)

	Next nX

End Transaction 

If cFilAnt <> cFilBkp	.AND. !Empty(cFilBkp)
	cFilAnt	:= cFilBkp

	OpenFile(cEmpAnt+cFilAnt)

EndIf 

If nErro > 0
    MsgAlert("Atenção!" + CRLF ;
            +"Ocorreram erros na criacao dos movimentos internos.";
            +"Total de erros: "+cValToChar(nErro), _cTitulo)


	If !ExistDir('\logs_SD3\')
		FWMakeDir('\logs_SD3\')
	EndIf 

	nHndl	:= FCreate('\logs_SD3\'+Dtos(dDataBase)+CriaTrab(,.F.)+".LOG")
	if nHndl = -1
        conout("Erro ao criar arquivo - ferror " + Str(Ferror()))
    else
        FWrite(nHndl, cMsgErr)
        FClose(nHndl)
    endif

    EECVIEW(cMsgErr)

Else
    MsgInfo("Processamento finalizado!") 

EndIf 

Return

Static Function OpenCSV(_cFile, aHeader, aDados)
Local cLine		:= ''

Local lPrim		:= .T. 

Local nHandle	:= 0
Local nLast		:= 0
Local nRecno	:= 0
Local nAux 		:= 0

Aadd(aHeader,'Filial'							 	)
Aadd(aHeader,'TP Movimento'					 		)
Aadd(aHeader,'Produto'					        	)
Aadd(aHeader,'Unidade'				    			)
Aadd(aHeader,'Quantidade'   						)
Aadd(aHeader,'Tipo RE/DE'	                    	)
Aadd(aHeader,'C Contabil'							)
Aadd(aHeader,'Armazem'								)
Aadd(aHeader,'Documento'    						)
Aadd(aHeader,'DT Emissao'       					)
Aadd(aHeader,'Grupo'        						)
Aadd(aHeader,'Custo'        						)
Aadd(aHeader,'Tipo'        							)
Aadd(aHeader,'Descr. Prod'  				    	)

nHandle := FT_FUse(_cFile)
If nHandle == -1  // Se houver erro de abertura abandona processamento
	MsgStop("Ocorreu um erro ao abrir o arquivo")

	Return
Endif

FT_FGoTop()// Posiciona na primeira linha

nLast := FT_FLastRec() // Retorna o número de linhas do arquivo
MsgAlert( cValtoChar( nLast ), 'Total de linhas' )

ProcRegua(nLast)
Do While !FT_FEOF()
	nAux ++ 
	
	IncProc( "Lendo o arquivo " + _cFile +  " | " + cValToChar( ROUND((nAux/nLast)*100,2))+ "%")

	cLine  := FT_FReadLn() // Retorna a linha corrente  
	nRecno := FT_FRecno()  // Retorna o recno da linha  

	If lPrim
		lPrim	:= .F. 
	Else 
		AADD(aDados,Separa(cLine,";",.T.))
	EndIf	

	FT_FSKIP() // Pula para próxima linha  
EndDo

FT_FUSE() // Fecha o arquivo

Return 

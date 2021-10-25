#Include "PROTHEUS.CH"
#Include "TOPCONN.CH"
#INCLUDE "RPTDEF.CH"
#INCLUDE "totvs.ch"
#INCLUDE "FILEIO.CH"

//variáveis do Log
Static cLogArquivo 	:= 'BaixaCSV'
Static cLogProcesso	:= ProcName()
Static cLogError	:= ''

user Function FINHAA01()
	local 	cMascara  	:= '*.csv'
	Local 	nMascpad 	:= 0
	local  	cDirini   	:= "\"
	Local  	lSalvar   	:= .T. //.T. = Salva || .F. = Abre
	Local  	nOpcoes   	:= GETF_LOCALHARD
	Local  	lArvore   	:= .T. //.T. = apresenta o árvore do servidor || .F. = não apresenta
	local   nSeq        := 1
	local   nDados      := 1

	Private     cLocalFile
	Private     aHead       := {}
	Private     cLeARQ      := ''
	Private     aDados      := {}

	cLocalFile	        := cGetFile( cMascara, "Escolha o arquivo", nMascpad, cDirIni, lSalvar, nOpcoes, lArvore)

	if !empty(cLocalFile)
		MV_PAR01 := cLocalFile
		cLeARQ	 := SubStr(MV_PAR01,1,RAT("\",MV_PAR01))
		FT_FUSE(MV_PAR01)
		FT_FGOTOP()

		While !FT_FEOF()
			if FT_FEOF()
				exit
			endif

			cLinha := Upper(FT_FREADLN())

			if nSeq = 1                             // cabeçalho

				aAdd(ahead,Separa(cLinha,";",.T.))
				ahead := ahead[1]

			else                                    // restante do CSV
				if Alltrim(cLinha) <> ""
					aAdd(aDados,Separa(cLinha,";",.T.))
					nDados++
				endif
			endIf

			FT_FSKIP()
			nSeq++
		EndDo

		Processa({|| BaixaSE2(aDados)}, "Baixando Títulos...")
	endIf
return

//-------------------------------------------------------------------------------------------------------------------------------------------------------
//                                                                      BaixaSE2
//-------------------------------------------------------------------------------------------------------------------------------------------------------

Static function  BaixaSE2(aDados) //01 001 000000001 A NF
	Local aBaixa 	:= {}
	Local cMotBx	:= getmv('FS_PAGMOTB',.T.,'NOR')
	Local cHist     := "FINHAA01 | " + dtoc(date()) + " | " + RetCodUsr() + " " + UsrFullName()
	Local dDat
	Local nLoop
	Local nJuros    := 0
	Local nDesc1    := 0
	Local nValor    := 10
	Local cFili     := ""
	Local cPrefixo
	Local cNum
	Local cParcela
	Local cTipo
	Local nI
	Local cMsgErro	:= ""
	Local lRet      := .T.
	Local cFilOld   := ''

	Private lMsErroAuto		:= .F. //Determina se houve algum tipo de erro durante a execucao do ExecAuto
	Private lMsHelpAuto		:= .T. //Define se mostra ou não os erros na tela (T= Nao mostra; F=Mostra)
	Private lAutoErrNoFile	:= .T. //Habilita a gravacao de erro da rotina automatica

	ProcRegua(Len(aDados))
	IncProc()

	for nI := 1 to Len(aDados)

		IncProc("Baixando registro " + cValToChar(nI) + " de " + cValToChar(Len(aDados)) + "...")

		cFili        := PadL(aDados[nI][1]       , TamSx3("E2_FILIAL")[1]    , "0")
		cPrefixo     := PadL(aDados[nI][2]       , TamSx3("E2_PREFIXO")[1]   , "0")
		cNum         := PadL(aDados[nI][3]       , TamSx3("E2_NUM")[1]       , "0")
		cParcelaB    := aDados[nI][4]

		if AllTrim(cParcelaB) == ""     //vazio ou número com zeros a esquerda
			cParcela     := PadR(cParcelaB        , TamSx3("E2_PARCELA")[1]   , " ")
		else
			cParcela     := PadL(cParcelaB        , TamSx3("E2_PARCELA")[1]   , "0")
		endif

		cTipo        := PadR(aDados[nI][5]       , TamSx3("E2_TIPO")[1]      , " ")
		cFornecedor  := PadL(aDados[nI][6]       , TamSx3("E2_FORNECE")[1]   , "0")
		cLoja        := PadL(aDados[nI][7]       , TamSx3("E2_LOJA")[1]   , "0")
		nValor       := Val(StrTran( aDados[nI][8], ",", "." ))
		cBanco       := aDados[nI][9]
		cAgencia     := aDados[nI][10]
		cConta       := aDados[nI][11]
		cMtBaixa     := aDados[nI][12]

		cFilAnt := cFili
		OpenFile(cEmpAnt+cFili)
/*
		If !(cFilAnt == cFili)
            cFilOld := cFilAnt
			IF RecLock("SM0",.F.)
                cFilAnt := cFili
                SM0->M0_CODFIL := cFilAnt
                SM0->(MsUnLock())
			Endif
		Endif
        cNumEmp := alltrim(cEmpAnt)+cFilAnt
*/
		dbSelectArea("SA6")
		SA6->(dbSetOrder(1))

		if SA6->(dbSeek(xFilial("SA6")+cBanco+cAgencia+cConta))
			cAgencia     := SA6->A6_AGENCIA
			cConta       := SA6->A6_NUMCON
		endif

		cMsgErro	+= CRLF + "Título:" + CRLF
		cMsgErro	+= "Filial: '" +cFili+ "' Prefixo: '" +cPrefixo+ "' Número: '" +cNum+ "' Parcela: '" +cParcela+ "' Tipo: '" +cTipo + "'" + " Fornecedor: '" +cFornecedor + "'" + " Loja: '" +cLoja + "'" + CRLF//fornecedor e loja

		dbSelectArea("SE2")
		SE2->(dbSetOrder(1))
		if SE2->(dbSeek(xFilial("SA6")+cPrefixo+cNum+cParcela+cTipo+cFornecedor+cLoja))

			dDataBase := CalcData(SE2->E2_EMISSAO)

			if nValor == SE2->E2_VALOR
				aBaixa :=   {}
				Aadd(aBaixa, {"E2_FILIAL"   ,SE2->E2_FILIAL              ,nil    })
				Aadd(aBaixa, {"E2_PREFIXO"  ,SE2->E2_PREFIXO             ,Nil    })
				Aadd(aBaixa, {"E2_NUM"      ,SE2->E2_NUM                 ,Nil    })
				Aadd(aBaixa, {"E2_PARCELA"  ,SE2->E2_PARCELA             ,Nil    })
				Aadd(aBaixa, {"E2_TIPO"     ,SE2->E2_TIPO                ,Nil    })
				Aadd(aBaixa, {"E2_FORNECE"  ,SE2->E2_FORNECE             ,Nil    })
				Aadd(aBaixa, {"E2_LOJA"     ,SE2->E2_LOJA                ,Nil    })
				Aadd(aBaixa, {"AUTMOTBX"    ,cMtBaixa                    ,Nil    })
				Aadd(aBaixa, {"AUTBANCO"    ,cBanco                      ,Nil    })
				Aadd(aBaixa, {"AUTAGENCIA"  ,cAgencia                    ,Nil    })
				Aadd(aBaixa, {"AUTCONTA"    ,cConta                      ,Nil    })
				Aadd(aBaixa, {"AUTDTBAIXA"  ,dDataBase                   ,Nil    }) //data da baixa
				Aadd(aBaixa, {"AUTDTCREDITO",dDataBase                   ,Nil    }) //data do credito
				Aadd(aBaixa, {"AUTHIST"     ,cHist        	             ,Nil    })
				Aadd(aBaixa, {"AUTVLRPG"    ,nValor                      ,NIL    })

				MSExecAuto({|x,y| Fina080(x,y)},aBaixa,3)

				If lMsErroAuto
					lRet        := .F.
					aErrPCAuto	:= GETAUTOGRLOG()
					cMsgErro	+= "ERRO! Problema no ExecAuto da baixa:" + CRLF
					For nLoop := 1 To Len(aErrPCAuto)
						cMsgErro += aErrPCAuto[nLoop] + CRLF
					Next
				else
					cMsgErro	+= "SUCESSO! Título baixado." + CRLF
				Endif
			else
				lRet        := .F.
				cMsgErro	+= "ERRO! Título não baixado. O valor informado no CSV: R$ "+AllTrim(TRANSFORM(nValor, "@E 999,999,999.99")) +" é diferente do valor do título gravado no E2_VALOR (contas a pagar): R$ "+AllTrim(TRANSFORM(SE2->E2_VALOR, "@E 999,999,999.99")) + CRLF + "Por favor verifique se o valor informado no CSV está correto."
			endif

		else

			cMsgErro	+= "ERRO! Título não encontrado na SE2 para ser baixado" + CRLF
			lRet        := .F.
		endif

	next

	LogFunction(cLogArquivo,cMsgErro,cLogProcesso)

	if lRet
		MsgInfo("Todos os títulos contidos no CSV foram baixados com sucesso!","Baixa de Títulos")
	else
		MsgInfo("Houve problema em pelo menos um dos títulos contidos no CSV. Para obter mais detalhes, por favor, acesse o arquivo de log que foi gravado no mesmo diretório do arquivo CSV de entada.","Baixa de Títulos")
	endif
    /*
	If !(cFilAnt == cFilOld)
        cFilAnt := cFilOld
		IF RecLock("SM0",.F.)
           
            SM0->M0_CODFIL := cFilAnt
             SM0->(MsUnLock())
		Endif
	Endif
    */
	dDataBase := date()
Return



//-------------------------------------------------------------------------------------------------------------------------------------------------------
//                                                                      LogFunction
//-------------------------------------------------------------------------------------------------------------------------------------------------------

Static Function LogFunction(cArquivo,cError,cProcesso,lMessage)
	Local cDiretorio	:= "\Log_Ethosx\"
	Local nHandle		:= 0
	Local cDir

	Default cArquivo 	:= "BaixaCSV.log"
	Default cError		:= "Log de erro"
	Default cProcesso	:= "Processo erro"
	Default lMessage	:= .F.

	If !('.log' $ cArquivo)
		cArquivo := cArquivo + '.log'
	Endif

	cArquivo 	:= cDiretorio+cArquivo

	If !(ExistDir(cDiretorio))
		MakeDir(cDiretorio)
	Endif

	if FILE(cArquivo) == .T.
		nHandle := fopen(cArquivo , FO_READWRITE + FO_SHARED )
		If nHandle != -1
			FSeek(nHandle, 0, FS_END)         // Posiciona no fim do arquivo
		Endif
	else
		nHandle := FCREATE(cArquivo)//Cria arquivo
	Endif

	If nHandle != -1
		If lMessage
			FWrite(nHandle,"[PROCCESS LOG:([" + AllTrim(Str(ThreadID())) + "]," + LogUserName() + "," + ComputerName() + ") " + Dtoc(date()) + "     " + GetRmtTime() + CRLF)
			FWrite(nHandle,"["+cValToChar(FWTimeStamp(3,,)+"]|")+ "["+cProcesso+"]|Log:"+cError + CRLF)
		Else
			FWrite(nHandle,"-------------------------------------------------------------------------------------------------------------------------"+CRLF+"[PROCCESS ERROR:([" + AllTrim(Str(ThreadID())) + "]," + LogUserName() + "," + ComputerName() + ") " + Dtoc(date()) + "     " + GetRmtTime() + CRLF)
			FWrite(nHandle,"["+cValToChar(FWTimeStamp(3,,)+"]|")+ "["+cProcesso+"]" + CRLF + cError + CRLF)
		Endif

		FClose(nHandle)
	Endif

	cDir := SubStr(cLocalFile, 1, Rat("\",cLocalFile))

	CpyS2T( cArquivo, cDir )
Return

//-------------------------------------------------------------------------------------------------------------------------------------------------------
//                                                                      CalcData
//-------------------------------------------------------------------------------------------------------------------------------------------------------

Static Function CalcData(dDataEmissão)
	Local dRet
	Local dDataFin := StoD("20201231")//StoD(SuperGetMv("MV_DATAFIN",.T.))

	if dDataEmissão <= dDataFin        //retornar 1o dia util do ano após a data
		dRet := LastDay(MonthSum(dDataFin,1),1)
	else
		dRet := dDataEmissão + 1
	endif

Return dRet

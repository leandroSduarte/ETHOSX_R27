#INCLUDE "PROTHEUS.CH"
#INCLUDE "TOPCONN.CH"

Static _cSemaforo := "HBARE250"
Static _lDtCorte  := .T.	//	Data de Corte Válida

//=====================================================================================
/*/{Protheus.doc} HBARE250
Serviços para Integração de Inclusão / Alteração Títulos a Receber Faturamento (Dif. Cupons)  
@author   Rafael Tenorio da Costa 
@history: Mário A. Cavenaghi - EthosX
@	24/05/21; Removido tratamento de existência do campo E1_XCHAVE e Não grava LOG se DOC dentro da Data de Corte
@	26/05/21; Verifica se o Tipo e/ou a Filial estão bloqueados
@	31/05/21; Inclusão de semáforo das tabelas Z96 e ZCA
@	04/06/21; Ajustes: No modo Console; Abertura das filiais liberadas; Flags de integração; Reenvio de inclusão
@	17/06/21; Removido E1_PARCELA do cChave e na gravação do E1_XCHAVE, porque interfere na localização da Baixa
/*/
User Function HBARE250(cEmpInt, cFilInt)
	Local cSelect := ""
	Local cFuncao := ""
	Local cError  := ""
	Local bError
	Local cTabZ90 := GetNextAlias()

	Private aStatSem  := {}
	Private aSemaforo := {}
	Private cHB_NUMRA := ' '	//SuperGetMv("HB_NUMRA"  ,.F.,"000000296") 19-02-2021
	Private cEX_RCGCMCT := ' '	//SuperGetMv("EX_RCGCMCT",.F.,"22749835")
	Private cTpOper := "10"	//	10 = Inclusão de título	//	29 = Alteração de título

	If IsBlind()
		Default cEmpInt := '01'
		Default cFilInt := '0001010001'

		RpcSetType(3)
		RpcSetEnv(cEmpInt, cFilInt)
	Else
		If cEmpInt == Nil
			cEmpInt := cEmpAnt
			cFilInt := cFilAnt
		EndIf
		If cEmpAnt <> cEmpInt
			cEmpAnt := cEmpInt
			cNumEmp := cEmpAnt + cFilAnt
			dbCloseAll()
			OpenSM0(cNumEmp)
			OpenFile(cNumEmp)
		Elseif cFilAnt <> cFilInt
			cFilAnt := cFilInt
			cNumEmp := cEmpAnt + cFilAnt
		Endif
	Endif

	bError := ErrorBlock({ |oError| cError := oError:Description + CRLF + oError:ErrorStack})
	cHB_NUMRA := SuperGetMv("HB_NUMRA"  ,.F.,"000000296")
	cEX_RCGCMCT := SuperGetMv("EX_RCGCMCT",.F.,"22749835")
	Z91->(dbSetOrder(1))	//	Z91_FILIAL + Z91_TPOPER + Z91_CHAVE + DTOS(Z91_DTOPER) + Z91_HROPER

	//=====================================================================================
	// Controle de semaforo - Nao permitir execução de mais de uma instância:
	//=====================================================================================
	cFuncao := PADR("HBARE250",TamSX3("Z97_ROTINA")[1])
	aStatSem := U_HBVERSEMAF(cFuncao)

	//Verifica se existe registro de semáforo   da rotina
	If !aStatSem[1]
		U_HBSMLog("SM",cFilAnt,cFuncao,"499","","","","Codigo chave da rotina "+cFuncao+" nao encontrada na tabela Z97.","499", aStatSem[2], .T.)
		Return
	EndIf

	//Verifica semáforo aberto
	If aStatSem[3] == "2"
		U_HBSMLog("SM",cFilAnt,aStatSem[4],"499","","","","Tentativa de execucao com Semaforo FECHADO","499", aStatSem[2], .T.)
		Return
	EndIf

	//Fecha o semáforo para execução
	aSemaforo := U_HBGRVSEMAF("2", aStatSem[2], aStatSem[4])
	If !aSemaforo[1]
		U_HBSMLog("SM",cFilAnt,aStatSem[4],"499","","","",aSemaforo[2],"499", aStatSem[2], .T.)
		Return
	EndIf
	//SetFunName("ARENAJOB250")	//	REMOVER quando atualizar a HBAREZ91

	U_LimpaZ91("SM",aStatSem[2], DtoS(Date()),'499')	//Apaga o controle de semáforo fechado

	U_HBJOB01A(cTabZ90, cTpOper)
	BEGIN SEQUENCE
		While !(cTabZ90)->( Eof() )
			cEmpAnt := (cTabZ90)->Z90_EMPRES
			cFilAnt := (cTabZ90)->Z90_FILIAL
			Conout("[HBARE250] "+Alltrim(cFilAnt)+"-"+Alltrim((cTabZ90)->Z90_AREFIL)+"-"+Alltrim((cTabZ90)->Z90_NOMFIL))
			MontaJson(cTabZ90)
			(cTabZ90)->( DbSkip() )
		EndDo
		RECOVER
		ConOut("[HBARE250] Erro Fatal")
	END SEQUENCE

	//Restaurando bloco de erro do sistema
	ErrorBlock(bError)

	//Se houve erro de processamento, será gravado na Z91
	If !Empty(cError)
		U_HBSMLog("SM",cFilAnt,"HBARE250","499","","",,"Processamento da rotina HBARE250 na filial "+AllTrim(cFilAnt)+" Interrompido. Vefique o LOG do serviço de JOB Schedule.","499", aStatSem[2], .T.)
		ConOut(" "+CRLF+CRLF)
		ConOut("[HBARE250] Houve um erro durante a transmissão na FILIAL: "+AllTrim(cFilAnt)+CRLF+cError)
		ConOut(" "+CRLF+CRLF)
	EndIf
	(cTabZ90)->(dbCloseArea())

	//Abre o semáforo após execução
	aSemaforo := U_HBGRVSEMAF("1", aStatSem[2], aStatSem[4])
	If !aSemaforo[1]
		U_HBSMLog("SM",cFilAnt,aStatSem[4],"499","","","",aSemaforo[2],"499", aStatSem[2], .T.)
	EndIf

	aSize(aStatSem, 0)
	aSize(aSemaforo, 0)

Return


//=====================================================================================
/*/{Protheus.doc} MontaJson
Faz a montagem do json que será enviado por filial
@author  Rafael Tenorio da Costa
@version 1.0
/*/
Static Function MontaJson(cTabZ90)
	Local cUrl := SuperGetMV("FS_AR25URL",, "http://ddnss-vmapp.ddns.com.br:7001/Arena/API_01/TSM/RecAltBxCan")
	Local cTabSE1 := GetNextAlias()
	Local cSelect := ""
	Local cChave := ""
	Local cJson := ""
	Local cJsonInc := ""
	Local cRetorno := ""
	Local cUrlCanc := ""
	Local aCmpAtu := {}
	Local lInclui := .F.
	Local lCancel := .F.
	Local lExclui := .F.
	Local nRegs := 0
	Local nRecno := 0
	Local cNotOrig := ""
	Local cNotOrg := ""
	Local cNotTipo := ""
	Local cNotTp := ""
	Local lLockE1 := .F.
	Local lIntegr := .F.
	Local aZ91 := {}
	Local nTop := SuperGetMv("FS_AREFIL",, 0)
	Local cTpOpCan := Alltrim(Posicione("Z96", 1, FWxFilial("Z96") + cTpOper, "Z96_CHAVEC"))	//	60 = Exclusão de título

	Private dDtCorte := Z96->Z96_DTFECH
	Private lSalto := .T.
	Private cTitulo := ""
	Private aRecnos := {}

	SA1->(dbSetOrder(1))	//	A1_FILIAL + A1_COD + A1_LOJA
	SA2->(dbSetOrder(1))	//	A2_FILIAL + A2_COD + A2_LOJA
	cNotOrig := Alltrim(GetNewPar("FS_ARE2502", "LOJA701"))
	cNotOrg  := "E1_ORIGEM NOT IN " + FormatIn(cNotOrig, ",")
	cNotTipo := Alltrim(GetNewPar("FS_ARE2503", "RA,NCC,AB,R$,CD,CC,VA"))
	cNotTp   := "E1_TIPO   NOT IN " + FormatIn(cNotTipo, ",")
	If nTop > 0	//	Filtra os 1os registros solicitados
		cSelect := " SELECT TOP " + cValToChar(nTop) + CRLF
	Else
		cSelect := " SELECT " + CRLF
	Endif
	cSelect +=    "        E1_FILIAL, E1_PREFIXO, E1_NUM, E1_TIPO, E1_PARCELA, E1_NUMLIQ, E1_XEXPTIT, E1_VENCREA, E1_VALOR, " + CRLF
	cSelect +=    "        E1_ORIGEM, E1_FILORIG, E1_EMISSAO, E1.R_E_C_N_O_ RECNOSE1, E1.D_E_L_E_T_ DELETSE1," + CRLF
	cSelect +=    "        E1_CLIENTE, E1_LOJA, E1_MSEXP, E1_NUMRA, E1_XCHAVE, " + CRLF
	cSelect +=    "        F2_FILIAL, F2_DOC, F2_SERIE, F2_CLIENTE, F2_LOJA, F2_CHVNFE, F2_ESPECIE, F2_TIPO," + CRLF
	cSelect +=    "        ED_XRDARE, ED_DESCRIC" + CRLF
	cSelect +=    "   FROM " + RetSqlName("SE1") + " E1" + CRLF
	cSelect +=    "   LEFT JOIN " + RetSqlName("SF2") + " F2 " + CRLF
	cSelect +=    "    ON  F2_FILIAL  = E1_FILORIG " + CRLF
	cSelect +=    "    AND F2_DOC     = E1_NUM " + CRLF
	cSelect +=    "    AND F2_SERIE   = E1_SERIE " + CRLF
	cSelect +=    "    AND F2_CLIENTE = E1_CLIENTE " + CRLF
	cSelect +=    "    AND F2_LOJA    = E1_LOJA " + CRLF
	cSelect +=    "    AND F2.D_E_L_E_T_ = ' ' " + CRLF
	cSelect +=    "   LEFT JOIN " + RetSqlName("SED") + " ED " + CRLF
	cSelect +=    "    ON  ED_FILIAL = '" + xFilial("SED", cFilAnt) + "' " + CRLF
	cSelect +=    "    AND ED_CODIGO = E1_NATUREZ " + CRLF
	cSelect +=    "    AND ED.D_E_L_E_T_ = ' '" + CRLF
	cSelect +=    "  WHERE ((E1.D_E_L_E_T_ = ' ' AND E1_XEXPTIT IN (' ', 'A')) OR (E1.D_E_L_E_T_ = '*' AND E1_XEXPTIT IN(' ','2')))" + CRLF
	cSelect += 	  "    AND E1_FILORIG = '" + cFilAnt + "'" + CRLF
	cSelect +=    "    AND E1_EMISSAO > '" + DtoS(dDtCorte) + "'" + CRLF
	cSelect +=    "    AND " + cNotOrg + CRLF
	cSelect +=    "    AND " + cNotTp + CRLF
	cSelect += 	  "    AND E1_MSEXP = ' ' " + CRLF
	//cSelect += 	  "    AND E1_NUM = '01042021A' " + CRLF	//	REMOVER
	cSelect +=    "  ORDER BY E1_FILORIG, E1_EMISSAO, E1_PREFIXO, E1_NUM, E1_TIPO, E1_PARCELA" + CRLF
	cSelect := ChangeQuery(cSelect)
	DbUseArea(.T., "TOPCONN", TcGenQry(,, cSelect), cTabSE1, .T., .F.)
	While !(cTabSE1)->(Eof())
		nRegs++
		nRecno := (cTabSE1)->RECNOSE1
		cUrlCanc := ""
		lCancel := .F.
		lSalto := .T.
		lInclui := .F.
		lExclui := .F.
		lIntegr := .F.
		cJson := ""
		cRetorno := ""
		aSize(aRecnos, 0)
		aSize(aCmpAtu, 0)

		//Campos atualizados quando retorno OK
		aAdd(aCmpAtu, {"E1_MSEXP"  , DtoS(dDataBase)})
		aAdd(aCmpAtu, {"E1_XEXPTIT", "2"})

		aZ91 := {}
		aAdd(aZ91, StoD((cTabSE1)->E1_EMISSAO))
		aAdd(aZ91, (cTabSE1)->E1_CLIENTE)
		aAdd(aZ91, (cTabSE1)->E1_NUM)
		aAdd(aZ91, (cTabSE1)->E1_VALOR)
		aAdd(aZ91, "SE1")
		aAdd(aZ91, "1")
		aAdd(aZ91, (cTabSE1)->(E1_FILIAL+E1_PREFIXO+E1_NUM+E1_PARCELA+E1_TIPO))
		//cChave := GetChvE1(cTabSE1, '10', nRecno, (cTabZ90)->Z90_AREFIL, .T., cTpOpCan)
		cChave := U_GetChSE1(cTabSE1, '10', nRecno, (cTabZ90)->Z90_AREFIL, .T., cTpOpCan)
		SE1->(dbGoto(nRecno))

		If (cTabSE1)->DELETSE1 == "*"		//	**	Exclusão	**
			lCancel := .T.
			aCmpAtu := {}
			aAdd(aCmpAtu, {"E1_MSEXP"  , DtoS(Date())})
			aAdd(aCmpAtu, {"E1_XEXPTIT", "3"})
			cTpOper  := "10"
			cUrlCanc := "Canc"
			cJson := '{'
			cJson += '"id": "' + cChave + '",'
			cJson += '"loja": "' + (cTabZ90)->Z90_AREFIL + '",'
			cJson += '"tipo": "' + cTpOper + '"'
			cJson += '}'
			aAdd(aRecnos, SE1->(Recno()))

		ElseIf SE1->E1_XEXPTIT == "A"		//	**	Alteração	**
			cTpOper := "29"
			cChave  +=  (cTabSE1)->E1_PARCELA
			cJson := '{'
			cJson += '"id": "' + cChave + '",'
			cJson += '"loja": "' + (cTabZ90)->Z90_AREFIL + '",'
			cJson += '"tipo": "' + cTpOper + '",'
			cJson += '"idtitulo": "' + cChave + '",'
			cJson += '"vencimento": "' + DtoS(SE1->E1_VENCREA) + '"'
			cJson += '}'
			aAdd(aRecnos, SE1->(Recno()))

		Else		//	**	Inclusão	**
			lInclui := .T.
			cJson := HB250INC(cTabSE1, cTabZ90, cChave)
		Endif

		//	Faz a comunicação e envia o json
		_lDtCorte := .T.
		If !lInclui .Or. dDtCorte < SE1->E1_EMISSAO
			lIntegr := HB250ENV(cJson, cUrl, Iif(lCancel, cUrlCanc, ""), @cRetorno)
			If !lIntegr
				If !lInclui	//	Se for manutenção e não existe o Título
					cJsonInc := HB250INC(cTabSE1, cTabZ90, cChave)	//	Gera JSon de inclusão
					lIntegr := HB250ENV(cJsonInc, cUrl, "", @cRetorno)	//	e envia o Título
					U_HbAreGrv(cTpOper, cChave, cJsonInc, cRetorno, aRecnos, aCmpAtu, aZ91, cFilAnt) //	Grava o log de processamento
					If lIntegr
						lIntegr := HB250ENV(cJson, cUrl, Iif(lCancel, cUrlCanc, ""), @cRetorno)	//	Reenvia a Manutenção
					Endif
				Endif
			Endif
		Else	//	Não envia se a EMISSÃO estiver dentro da nova data de corte
			lExclui := .T.
		Endif

		//	Grava o log de processamento
		If _lDtCorte
			U_HbAreGrv(Iif(lCancel, cTpOpCan, cTpOper), cChave, cJson, cRetorno, aRecnos, aCmpAtu, aZ91, cFilAnt)
		Endif

		Conout("[HBARE250] "+Alltrim(Str(nRegs))+'."'+Iif(lCancel, cTpOpCan, cTpOper)+'" '+DTOC(dDataBase)+"-"+Time()+" "+cEmpAnt+"."+(cTabZ90)->Z90_FILIAL+"."+Alltrim((cTabZ90)->Z90_NOMFIL)+" Chave => "+Alltrim(cChave))
		If lIntegr
			U_LimpaZ91(cTpOper , cChave, DtoS(Date()), '401')
			U_LimpaZ91(cTpOpCan, cChave, DtoS(Date()), '401')
		Else
			Conout("[HBARE250] " + cRetorno)
		Endif
		If lIntegr .Or. lExclui
			For nRecno := 1 To Len(aRecnos)
				SE1->(dbGoTo(aRecnos[nRecno]))
				lLockE1 := .F.
				If lIntegr .And. Empty(SE1->E1_XCHAVE)
					SE1->(Reclock("SE1"))
					SE1->E1_XCHAVE := cTitulo
					lLockE1 := .T.
				Endif
				If Empty(SE1->E1_XEXPTIT)
					If !lLockE1
						SE1->(Reclock("SE1"))
						lLockE1 := .T.
					Endif
					SE1->E1_XEXPTIT := '3'
				Endif
				If lLockE1
					If Empty(SE1->E1_MSEXP)
						SE1->E1_MSEXP := DtoS(dDataBase)
					EndiF
					SE1->(msUnLock())
				Endif
			Next
		Endif
		If lSalto
			(cTabSE1)->(dbSkip())
		Endif
	Enddo
	If Z96->Z96_DTFECH <> dDtCorte
		Z96->(RecLock("Z96"))
		Z96->Z96_DTFECH := dDtCorte
		Z96->(msUnLock())
	Endif
	(cTabSE1)->(dbCloseArea())

Return(nRegs)


//=====================================================================================
//	Inclusão de Título
Static Function HB250INC(cTabSE1, cTabZ90, cChave)
	Local nVlrTotal := 0
	Local cNumNF := ""
	Local cRazao := ""
	Local cFantasia := ""
	Local cCNPJ := ""
	Local cContaSB1 := ""
	Local cDescSB1 := ""
	Local cChaveSF2 := ""
	Local cContaSED := ""
	Local cDescSED := ""
	Local cTipoSE1 := ""

	cTipoSE1 := AllTrim((cTabSE1)->E1_TIPO)
	cTpOper := "10"
	cRazao := ""
	cFantasia := ""
	cCNPJ := ""
	If AllTrim((cTabSE1)->F2_TIPO) == "B"
		If SA2->(dbSeek(xFilial("SA2",cFilAnt)+(cTabSE1)->E1_CLIENTE+(cTabSE1)->E1_LOJA))
			cRazao    := AllTrim(SA2->A2_NOME)
			cFantasia := AllTrim(SA2->A2_NREDUZ)
			cCNPJ     := AllTrim(SA2->A2_CGC)
		EndIf
	Else
		If SA1->(dbSeek(xFilial("SA1",cFilAnt)+(cTabSE1)->E1_CLIENTE+(cTabSE1)->E1_LOJA))
			cRazao    := AllTrim(SA1->A1_NOME)
			cFantasia := AllTrim(SA1->A1_NREDUZ)
			cCNPJ     := AllTrim(SA1->A1_CGC)
		EndIf
	EndIf

	//Alterado em 22/12/2020 - par gerar títulos aglutinados
	cNumNF := (cTabSE1)->E1_NUM + Iif(!Empty((cTabSE1)->E1_NUMLIQ), "-AGL", "")

	//SUPPLY-MERCATO - construir um parâmetro para cada situação
	If Alltrim((cTabSE1)->E1_ORIGEM) = "GSFINFT1"  .And. Alltrim((cTabSE1)->E1_TIPO) = "NF" //aglutinação na Supply
		If (cTabSE1)->E1_FILIAL = '2501' .And. Left(cCNPJ,8) = Alltrim(cEX_RCGCMCT) .And. (cTabSE1)->E1_PREFIXO = 'MCT' .And. Val(Alltrim((cTabSE1)->E1_NUM)) = Val(Alltrim((cTabSE1)->E1_NUMRA))
			cNumNF := (cTabSE1)->E1_NUM +"-AGL"
		Endif
	Endif

	//INCLUIR PROCESSO DO LEITE
	If Alltrim(cEmpAnt) = '03' .And. Alltrim((cTabSE1)->E1_PREFIXO) $ '|HAB|RAG|REX|POS|HEX|' //aglutinação da PPM
		If  Alltrim((cTabSE1)->E1_TIPO) = 'FT' .And. Alltrim((cTabSE1)->E1_ORIGEM) = 'FINA040'
			cNumNF := (cTabSE1)->E1_NUM +"-AGL"
		Endif
	Endif

	cTitulo := (cTabSE1)->(E1_FILORIG + E1_CLIENTE + E1_LOJA + E1_PREFIXO + E1_NUM + E1_TIPO)
	cChaveSF2 := (cTabSE1)->(F2_FILIAL + F2_DOC + F2_SERIE + F2_CLIENTE + F2_LOJA)
	cContaSED := (cTabSE1)->ED_XRDARE
	cDescSED  := AllTrim((cTabSE1)->ED_DESCRIC)
	cJson := '{'
	cJson += '"id": "' + cTitulo + '",'
	cJson += '"loja": "' + (cTabZ90)->Z90_AREFIL + '",'
	cJson += '"tipo": "' + cTpOper + Iif(cTipoSE1 $ "IR-|CF-|PI-|CS-", "-I", "")  + '",'
	cJson += '"dt_emissao": "' + (cTabSE1)->E1_EMISSAO + '",'
	cJson += '"num_nf": "' + cNumNF + '",'
	cJson += '"cnpjcli": "' + cCNPJ + '",'
	cJson += '"razao_cli": "' + cRazao + '",'
	cJson += '"fantasia_cli": "' + cFantasia + '",'
	If AllTrim((cTabSE1)->F2_ESPECIE) == "RPS"
		cJson += '"chv_nfe": "' + "SEM CHAVE" + '",'
	Else
		cJson += '"chv_nfe": "' + Iif(AllTrim((cTabSE1)->E1_ORIGEM) == "FINA040", "SEM CHAVE", (cTabSE1)->F2_CHVNFE) + '",'
	EndIf
	cJson += '"link_xml": "' + 'link.xml' + '",'            //esta TAG será analisada posteriormente, por causa de um problema no TSS
	cJson += '"historico": "' + cFantasia + ' ' + (cTabSE1)->E1_EMISSAO + ' ' + (cTabSE1)->E1_NUM + '",'
	cJson += '"titulo": ['
	lSalto := .T.
	While cTitulo == (cTabSE1)->(E1_FILORIG + E1_CLIENTE + E1_LOJA + E1_PREFIXO + E1_NUM + E1_TIPO) .And. !(cTabSE1)->(Eof())
		SE1->(dbGoto((cTabSE1)->RECNOSE1))
		cJson += '{'
		cJson += '"idtit": "' + cTitulo + (cTabSE1)->E1_PARCELA + '",'
		cJson += '"vencimento": "' + (cTabSE1)->E1_VENCREA + '",'
		cJson += '"valor": ' + cValToChar((cTabSE1)->E1_VALOR) + ','
		cJson += '"parcela": "' + (cTabSE1)->E1_PARCELA + '"'
		cJson += '},'
		nVlrTotal += (cTabSE1)->E1_VALOR
		aAdd(aRecnos, (cTabSE1)->RECNOSE1)
		(cTabSE1)->(dbSkip())
		lSalto := .F.
	EndDo
	cJson := SubStr(cJson, 1, Len(cJson)-1)
	cJson += '],'
	cJson += '"itens": ['
	SB1->(dbSetOrder(1))	//	B1_FILIAL + B1_COD
	SF4->(dbSetOrder(1))	//	F4_FILIAL + F4_CODIGO
	SD2->(dbSetOrder(3))	//	D2_FILIAL + D2_DOC + D2_SERIE + D2_CLIENTE + D2_LOJA + D2_COD + D2_ITEM
	If !Empty(cChaveSF2) .And. SD2->(dbSeek(cChaveSF2))
		nVlrTotal := 0
		While !SD2->(Eof()) .And. cChaveSF2 == SD2->(D2_FILIAL + D2_DOC + D2_SERIE + D2_CLIENTE + D2_LOJA)
			cContaSB1 := ""
			cDescSB1  := ""
			SF4->(dbSeek(xFilial("SF4", cFilAnt)+SD2->D2_TES))
			If AllTrim(SD2->D2_TIPO) $ 'P*I'
				nVlrTotal += SD2->D2_TOTAL + SD2->D2_VALFRE + SD2->D2_SEGURO + SD2->D2_DESPESA
			Else
				nVlrTotal += SD2->D2_TOTAL + SD2->D2_VALIPI + Iif(SF4->F4_INCSOL == "N", 0, SD2->D2_ICMSRET) + SD2->D2_VALFRE + SD2->D2_SEGURO + SD2->D2_DESPESA
			EndIf
			If SB1->(dbSeek(xFilial("SB1",cFilAnt) + SD2->D2_COD))
				cContaSB1 := Iif(SB1->(FieldPos("B1_XRRVEN")) > 0 .And. !Empty(SB1->B1_XRRVEN),SB1->B1_XRRVEN,SB1->B1_XRDARE)	//	SB1->B1_XCONTA
				cDescSB1  := AllTrim(SB1->B1_DESC)
			EndIf
			cJson += '{'
			cJson += '"id_item": "' + cTitulo + SD2->D2_ITEM + '",'	//	25/02/2021, Mário Cavenaghi, EthosX
			cJson += '"id_rd": "' + cContaSB1 + '",'
			If AllTrim(SD2->D2_TIPO) $ 'P*I'
				cJson += '"valor_item": ' + cValToChar(SD2->D2_TOTAL + SD2->D2_VALFRE + SD2->D2_SEGURO + SD2->D2_DESPESA) + ','
			Else
				cJson += '"valor_item": ' + cValToChar(SD2->D2_TOTAL + SD2->D2_VALIPI + Iif(SF4->F4_INCSOL == "N", 0, SD2->D2_ICMSRET) + SD2->D2_VALFRE + SD2->D2_SEGURO + SD2->D2_DESPESA) + ','
			EndIf
			cJson += '"historico_it": "' + cDescSB1 + '",'
			cJson += '"ccusto": "",'
			cJson += '"nomeccusto": ""'
			cJson += '},'
			SD2->(dbSkip())
		EndDo
		cJson := SubStr(cJson, 1, Len(cJson)-1)
	Else
		cJson += '{'
		cJson += '"id_item": "' + cChave + '",'
		cJson += '"id_rd": "' + cContaSED + '",'
		cJson += '"valor_item": ' + cValToChar(nVlrTotal) + ','
		cJson += '"historico_it": "' + cDescSED + '",'
		cJson += '"ccusto": "",'
		cJson += '"nomeccusto": ""'
		cJson += '}'
	EndIf
	cJson += '],'
	cJson += '"valor": ' + cValToChar(nVlrTotal)
	cJson += '}'

Return(cJson)


//=====================================================================================
//	Envia o JSON
Static Function HB250ENV(cJson, cUrl, cUrlCanc, cRetorno)
	Local lRet := .F.

	cJson := FwNoAccent(cJson)
	cRetorno := U_HbAreEnv(cJson, cUrl + cUrlCanc, /*cUsuario*/, /*cSenha*/, _cSemaforo)
	If '"401"' $ cRetorno
		If "DATA FORA DO PER" $ Upper(cRetorno)	//	Dentro da Data de Corte
			If dDtCorte <  SE1->E1_EMISSAO
				dDtCorte := SE1->E1_EMISSAO
				_lDtCorte := .F.	//	Não grava LOG
			Endif
			lSalto := .T.	//	Se também for uma exclusão, não irá enviar
			//Elseif "NÃO EXISTEM LANÇAMENTOS" $ Upper(cRetorno)
			//HB250INC(cTabSE1, cTabZ90, cChave)
		Endif
	ElseIf '"200"' $ cRetorno
		lRet := .T.
	Endif

Return(lRet)


//=====================================================================================
//	Verifica se o Título já foi enviado e qual formato de chave
User Function GetChSE1(cTabSE1, cTpOper, nRecno, cAreFil, lLimpa, cTpOpCan)
	Local cChave1 := (cTabSE1)->(E1_FILORIG + E1_EMISSAO + E1_PREFIXO + E1_NUM + E1_TIPO)// 1a. Chave, sem parcela. Não é chave única no sistema
	Local cChave2 := RetSqlName("SE1") + STRZERO(nRecno, 10) + cAreFil + (cTabSE1)->( E1_CLIENTE + E1_PREFIXO + E1_NUM + E1_TIPO + E1_PARCELA)
	Local cChave3 := (cTabSE1)->(E1_FILORIG + E1_CLIENTE + E1_LOJA + E1_PREFIXO + E1_NUM + E1_TIPO)	//	Atual Chave Única
	Local cChave  := ""

	If !Empty((cTabSE1)->E1_XCHAVE)
		If At('|', (cTabSE1)->E1_XCHAVE) > 0
			cChave := StrTran((AllTrim((cTabSE1)->E1_XCHAVE)), '|')
		Else
			cChave := (cTabSE1)->E1_XCHAVE
		Endif
	ElseIf U_fLerZ91(cFilAnt, cChave1, cTpOper, "200")
		cChave := cChave1
	ElseIf U_fLerZ91(cFilAnt, cChave2, cTpOper, "200")
		cChave := cChave2
	Else
		cChave := cChave3
		If lLimpa
			U_LimpaZ91(cTpOper, cChave1, DtoS(Date()), '401',, cFilAnt)
			U_LimpaZ91(cTpOper, cChave2, DtoS(Date()), '401',, cFilAnt)
			If !Empty(cTpOpCan)
				U_LimpaZ91(cTpOpCan, cChave1, DtoS(Date()), '401',, cFilAnt)
				U_LimpaZ91(cTpOpCan, cChave2, DtoS(Date()), '401',, cFilAnt)
			EndIf
		EndIf
	EndIf

Return(cChave)

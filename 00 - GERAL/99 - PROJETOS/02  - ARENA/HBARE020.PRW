#INCLUDE "TOPCONN.CH"
#INCLUDE "PROTHEUS.CH"

#DEFINE TPOPERACAO "24"

Static _aUF := {}

/*/{Protheus.doc} HBARE020
Rotina de Integração de Cadastro de Funcionários com o ARENA
@type class
@author   Edson Melo
@since    27/04/20
@history  
@ Mário Cavenaghi, EthosX
@	20/05/21: Inclusão da Data de Corte do campo Z96_DTFECH
@	21/05/21: Query revisada removendo o UNION e aplicado no campo CANC da Query
@	22/05/21: Utilização da tabela '12' do SX5 para tratar estados
@	26/05/21: Tratamento para rodar como JOB ou Console
@	28/05/21: Incluido número de repetições para Reenvio de registro com TimeOut
@	18/05/21: Não enviará mais se o Cancelamento der erro '401' com Data antes do Corte
@	20/08/21: Incluído Z90_FILIAL da Query que retorna a LOJA do ARENA
/*/
User Function HBARE020(aEmpFil)
	Local cAliasZ90

	Private aEmpFilZ90 := {}
	Private cFilZ90 := ""

	If ValType(aEmpFil) <> 'A'
		aEmpFil := {'01', '0001010001'}
	Endif
	If IsBlind()
		RpcSetType(3)
		RpcSetEnv(aEmpFil[1], aEmpFil[2])
	ElseIf cEmpAnt <> aEmpFil[1]
		cEmpAnt := aEmpFil[1]
		cNumEmp := cEmpAnt + cFilAnt
		dbCloseAll()
		OpenSM0(cNumEmp)
		OpenFile(cNumEmp)
	Elseif cFilAnt <> aEmpFil[2]
		cFilAnt := aEmpFil[2]
		cNumEmp := cEmpAnt + cFilAnt
	Endif

	// Verifica se existe registro de semáforo da rotina
	cFuncao  := PADR("HBARE020", TamSX3("Z97_ROTINA")[1])
	aStatSem	:= U_HBVERSEMAF(cFuncao)
	If !aStatSem[1]
		U_HBSMLog("SM",cFilAnt,cFuncao,"499","","","","Codigo chave da rotina "+cFuncao+" nao encontrada na tabela Z97.","499", aStatSem[2],.t.)
		Return
	EndIf

	// Verifica semáforo aberto
	If aStatSem[3] == "2"
		U_HBSMLog("SM",cFilAnt,aStatSem[4],"499","","","","Tentativa de execucao com Semaforo FECHADO","499", aStatSem[2], .t.)
		Return
	EndIf

	// Fecha o semáforo para execução
	aSemaforo := U_HBGRVSEMAF("2", aStatSem[2], aStatSem[4])
	If !aSemaforo[1]
		U_HBSMLog("SM",cFilAnt,aStatSem[4],"499","","","",aSemaforo[2],"499", aStatSem[2], .t.)
		Return
	EndIf

	// Apaga o controle de semáforo fechado
	U_LimpaZ91("SM",aStatSem[2], DtoS(Date()),'499')

	cAliasZ90 := GetNextAlias()
	U_HBJOB01A(cAliasZ90, TPOPERACAO)
	While (cAliasZ90)->(!Eof())
		cEmpZ90 := (cAliasZ90)->Z90_EMPRES
		cFilZ90 := (cAliasZ90)->Z90_FILIAL
		aAdd(aEmpFilZ90,{cEmpZ90,cFilZ90})
		(cAliasZ90)->(dbSkip())
	Enddo
	(cAliasZ90)->(dbCloseArea())
	If Len(aEmpFilZ90) > 0
		fJobArena(aEmpFilZ90)
	EndIf
	aSemaforo := U_HBGRVSEMAF("1", aStatSem[2], aStatSem[4])
	If !aSemaforo[1]
		U_HBSMLog("SM",cFilAnt,aStatSem[4],"499","","","",aSemaforo[2],"499", aStatSem[2], .T.)
	EndIf
	aSize(aStatSem,0)
	aSize(aSemaforo,0)

Return


/*/{Protheus.doc} fJobArena
@author Robson B. Melo
@since Abr/2020
@version 1.00
/*/
Static Function fJobArena(aEmpFilZ90)
	Local nX

	Private cEmpZ90 := ""
	Private cFilZ90 := ""
	Private cError := ""
	Private bError := ErrorBlock({ |oError| cError := oError:Description+CRLF+oError:ErrorStack})

	Conout("[HBARE020] - Início: " + DtoC(dDataBase) + " - " + Time())
	For nX := 1 to Len(aEmpFilZ90)
		cEmpZ90 := aEmpFilZ90[nX, 1]
		cFilZ90 := aEmpFilZ90[nX, 2]
		cFilAnt := cFilZ90
		BEGIN SEQUENCE
			ProcSRA()
		END SEQUENCE
		ErrorBlock(bError)

		//Se houve erro de processamento, será gravado na Z91
		If !Empty(cError)
			U_HBSMLog("SM",cFilAnt,"HBARE020","499","","",,"Processamento da rotina HBARE020 na filial "+Alltrim(cFilZ90)+" Interrompido. Vefique o LOG do serviço de JOB Schedule.","499", "0030", .T.)
			ConOut("[HBARE020] - Houve um erro durante a transmissão na FILIAL: "+Alltrim(cFilZ90)+CRLF+cError)
		EndIf
	Next
	Conout("[HBARE020] - Fim: " + DtoC(dDataBase) + " - " + Time())

Return


/*/{Protheus.doc} ProcSRA
Rotina responsável por processar todos os registros não enviados
@author Edson Melo
@since 2704/2020
@type class
/*/
Static Function ProcSRA()
	Local aHeader     := {}
	Local cTmpAlias   := GetNextAlias()
	Local cUrl        := GetNewPar("FS_FOLURL" , "")
	Local cPath       := GetNewPar("FS_PATHFUN", "")
	Local cPathCan    := GetNewPar("FS_CANFUN" , "")
	Local oRestClient := FWRest():New(cUrl)
	Local cJsonResp   := ""
	Local cUser       := Alltrim(GetNewPar("FS_USRAREN", ""))
	Local cPassword   := Alltrim(GetNewPar("FS_PASWARE", ""))
	Local cTpInt120   := GetMV("FS_120TPOP")
	Local cTpCanc     := Alltrim(Posicione("Z96", 1, xFilial("Z96") + cTpInt120, "Z96_CHAVEC"))
	Local cDtCorte    := DtoS(Z96->Z96_DTFECH)
	Local nPosUF      := 0
	Local cDescUF     := ""
	Local cSelect     := ""
	Local aInclusao   := {}
	Local aPath       := {}
	Local lIntegr     := .F.
	Local cCodErr     := ""
	Local cCodAmb     := ""
	Local oJSon       := Nil
	Local nTimeOut    := 2
	Local nTime       := 0
	Local lCorte      := .F.

	If Empty(_aUF)
		SX5->(dbSetOrder(1))
		SX5->(dbSeek(xFilial() + '12'))	//	Carrega a Tabela de Estados para o Array
		While !SX5->(Eof()) .And. xFilial("SX5") + '12' == SX5->X5_FILIAL + SX5->X5_TABELA
			aAdd(_aUF, {Alltrim(SX5->X5_CHAVE), Alltrim(SX5->X5_DESCRI)})
			SX5->(dbSkip())
		Enddo
	Endif

	If Empty(cUrl)
		Contou('[HBARE110] Url não definida.')
		Return
	Endif

	//aAdd(aHeader, "Content-Type: application/json")
	aAdd(aHeader, "Content-Type:application/json; Charset=UTF-8")
	aAdd(aHeader, 'Authorization: Basic ' + Alltrim(OemToAnsi(Encode64(cUser + ":" + cPassword))))

	Conout("[HBARE020] - " + cEmpAnt + "." + cFilAnt)
	BeginSQL Alias cTmpAlias
		COLUMN RECSRA as Numeric(10, 0)
		SELECT 
			(CASE WHEN SRA.D_E_L_E_T_ = ' ' THEN 'N' ELSE 'S' END) CANC,
			RA_FILIAL+RA_MAT ID, RA_FILIAL, RA_CIC, RA_RG, RA_NOME, RA_APELIDO, RA_ENDEREC, RA_COMPLEM, RA_BAIRRO, RA_MUNICIP, RA_ESTADO,
			RA_CEP, RA_NASC, RA_CC, ISNULL(CTT_DESC01, '') CTT_DESC01, RA_DEPTO, ISNULL(QB_DESCRIC, '') QB_DESCRIC, RA_TNOTRAB, 
			ISNULL(R6_DESC, '') R6_DESC, RA_SEQTURN, RA_XDFOLGA, RJ_CODCBO, RJ_DESC, RA_CODFUNC, RA_MAT, RA_ADMISSA, RA_SEXO, 
			(CASE WHEN RA_SEXO = 'M' THEN 'MASCULINO' ELSE (CASE WHEN RA_SEXO = 'F' THEN 'FEMININO' ELSE '' END) END) SEXO, RA_SITFOLH, 
			RA_SALARIO, Q3_DESCSUM, RA_XTREINA, RA_TELEFON, RA_EMAIL, RA_ESTCIVI, RA_NUMCP, RA_DEMISSA, RJ_XAREFUN, 	RJ_XDESARE,
			( SELECT TOP 1 Z90_AREEMP+Z90_AREFIL AREEMP
				FROM   %Table:Z90% Z90
				WHERE	 Z90_FILIAL NOT IN
				(	SELECT ZCA_CODFIL 
					FROM   %table:ZCA% ZCA 
					WHERE  ZCA.%notDel%
					AND ZCA_FILIAL = %Exp:FWxFilial('ZCA',cFilZ90)%
					AND ZCA_CHAVE  = '24'
					AND ZCA_STATUS = '2'	// Desabilitado
				)
				AND Z90_EMPRES = %exp:cEmpZ90%
				AND Z90_FILIAL = %Exp:FWxFilial('Z90',cFilZ90)%
				AND Z90.%NotDel%
			) LOJA,
			SRA.R_E_C_N_O_ RECSRA
		FROM %Table:SRA% SRA
		LEFT JOIN %Table:CTT% CTT 
			ON	 CTT_FILIAL = %Exp:FWxFilial("CTT",cFilZ90)%
			AND CTT_CUSTO  = RA_CC
			AND CTT.D_E_L_E_T_  = ' '
		LEFT JOIN %Table:SQB% SQB 
			ON  QB_FILIAL = %Exp:FWxFilial("SQB",cFilZ90)%
			AND QB_DEPTO  = RA_DEPTO
			AND SQB.%NotDel%
		LEFT JOIN %Table:SR6% SR6 
			ON  R6_FILIAL = %Exp:FWxFilial("SR6",cFilZ90)%
			AND R6_TURNO  = RA_TNOTRAB
			AND SR6.%NotDel%
		LEFT JOIN %Table:SRJ% SRJ 
			ON  RJ_FILIAL = %Exp:FWxFilial("SRJ",cFilZ90)%
			AND RJ_FUNCAO = RA_CODFUNC
			AND SRJ.%NotDel%
		LEFT JOIN %Table:SQ3% SQ3 
			ON  Q3_FILIAL = %Exp:FWxFilial("SQ3",cFilZ90)%
			AND Q3_CARGO  = RA_CARGO
			AND SQ3.%NotDel%
		WHERE	 RA_FILIAL = %Exp:FWxFilial("SRA",cFilZ90)%
			AND RA_MSEXP  = ' ' 
		ORDER BY ID
	EndSQL

	BEGIN SEQUENCE
		While (cTmpAlias)->(!Eof())
			cTpInt120 := GetMV("FS_120TPOP")
			lCancel := .F.
			nPosUF  := aScan(_aUF, {|x| x[1] == (cTmpAlias)->RA_ESTADO})
			cDescUF := ""
			If nPosUF > 0
				cDescUF := _aUF[nPosUF, 2]
			Else
				cDescUF := Alltrim(Posicione('SX5', 1, xFilial("SX5", cFilZ90) + '12' + (cTmpAlias)->RA_ESTADO, 'X5_DESCRI'))
			Endif
			aDados := {;	//
			Iif(Empty((cTmpAlias)->ID),         "NAO INFORMADO", Alltrim((cTmpAlias)->ID)),;	//	01
			Iif(Empty((cTmpAlias)->LOJA),       "NAO INFORMADO", Alltrim((cTmpAlias)->LOJA)),;	//02
			Iif(Empty((cTmpAlias)->RA_CIC),     "NAO INFORMADO", Alltrim((cTmpAlias)->RA_CIC)),;	//03
			Iif(Empty((cTmpAlias)->RA_RG),      "NAO INFORMADO", Alltrim((cTmpAlias)->RA_RG)),;	//04
			Iif(Empty((cTmpAlias)->RA_NOME),    "NAO INFORMADO", Alltrim((cTmpAlias)->RA_NOME)),;	//05
			Iif(Empty((cTmpAlias)->RA_APELIDO), "NAO INFORMADO", Alltrim((cTmpAlias)->RA_APELIDO)),;	//06
			Iif(Empty((cTmpAlias)->RA_ENDEREC), "NAO INFORMADO", Alltrim((cTmpAlias)->RA_ENDEREC)),;	//07
			Iif(Empty((cTmpAlias)->RA_COMPLEM), "NAO INFORMADO", Alltrim((cTmpAlias)->RA_COMPLEM)),;	//08
			Iif(Empty((cTmpAlias)->RA_BAIRRO),  "NAO INFORMADO", Alltrim((cTmpAlias)->RA_BAIRRO)),;	 //09
			Iif(Empty((cTmpAlias)->RA_MUNICIP), "NAO INFORMADO", Alltrim((cTmpAlias)->RA_MUNICIP)),;	//10
			Iif(Empty((cTmpAlias)->RA_ESTADO),  "NAO INFORMADO", Alltrim((cTmpAlias)->RA_ESTADO)),;	 //11
			cDescUF,;	//12
			Iif(Empty((cTmpAlias)->RA_CEP),     "NAO INFORMADO", Alltrim((cTmpAlias)->RA_CEP)),;	//13
			Iif(Empty((cTmpAlias)->RA_NASC),    "NAO INFORMADO", Alltrim((cTmpAlias)->RA_NASC)),;	//14
			Iif(Empty((cTmpAlias)->RA_CC),      "NAO INFORMADO", Alltrim((cTmpAlias)->RA_CC)),;	//15
			Iif(Empty((cTmpAlias)->CTT_DESC01), "NAO INFORMADO", Alltrim((cTmpAlias)->CTT_DESC01)),;	//16
			Iif(Empty((cTmpAlias)->RA_DEPTO),   "NAO INFORMADO", Alltrim((cTmpAlias)->RA_DEPTO)),;	//17
			Iif(Empty((cTmpAlias)->QB_DESCRIC), "NAO INFORMADO", Alltrim((cTmpAlias)->QB_DESCRIC)),;	//18
			Iif(Empty((cTmpAlias)->RA_TNOTRAB), "NAO INFORMADO", Alltrim((cTmpAlias)->RA_TNOTRAB)),;	//19
			Iif(Empty((cTmpAlias)->R6_DESC),    "NAO INFORMADO", Alltrim((cTmpAlias)->R6_DESC)),;	//20
			Iif(Empty((cTmpAlias)->RA_SEQTURN), "NAO INFORMADO", Alltrim((cTmpAlias)->RA_SEQTURN)),;	//21
			Iif(Empty((cTmpAlias)->RA_XDFOLGA), "NAO INFORMADO", Alltrim((cTmpAlias)->RA_XDFOLGA)),;	//22
			Iif(Empty((cTmpAlias)->RJ_CODCBO),  "NAO INFORMADO", Alltrim((cTmpAlias)->RJ_CODCBO)),;	//23
			Iif(Empty((cTmpAlias)->RJ_DESC),    "NAO INFORMADO", Alltrim((cTmpAlias)->RJ_DESC)),;	//24
			Iif(Empty((cTmpAlias)->RA_CODFUNC), "NAO INFORMADO", Alltrim((cTmpAlias)->RA_CODFUNC)),;	//25
			Iif(Empty((cTmpAlias)->RJ_DESC),    "NAO INFORMADO", Alltrim((cTmpAlias)->RJ_DESC)),;	//26
			Iif(Empty((cTmpAlias)->RA_MAT),     "NAO INFORMADO", Alltrim((cTmpAlias)->RA_MAT)),;	//27
			Iif(Empty((cTmpAlias)->RA_ADMISSA), "NAO INFORMADO", Alltrim((cTmpAlias)->RA_ADMISSA)),;	//28
			Iif(Empty((cTmpAlias)->RA_SEXO),    "NAO INFORMADO", Iif((cTmpAlias)->RA_SEXO == 'M', 'MASCULINO', 'FEMININO')),;	//29
			Iif(Empty((cTmpAlias)->SEXO),       "NAO INFORMADO", Alltrim((cTmpAlias)->SEXO)),;	//30
			Iif(Empty((cTmpAlias)->RA_SITFOLH), "NAO INFORMADO", Alltrim((cTmpAlias)->RA_SITFOLH)),;	//31
			Iif(Empty((cTmpAlias)->RA_SITFOLH), "NAO INFORMADO", Alltrim(TABELA("31", (cTmpAlias)->RA_SITFOLH))),;	//32
			Iif(Empty((cTmpAlias)->RA_SALARIO), "NAO INFORMADO", Alltrim(Str((cTmpAlias)->RA_SALARIO))),;	//33
			Iif(Empty((cTmpAlias)->Q3_DESCSUM), "NAO INFORMADO", Alltrim((cTmpAlias)->Q3_DESCSUM)),;	//34
			Iif(Alltrim((cTmpAlias)->RA_XTREINA) == "", "N"    , Alltrim((cTmpAlias)->RA_XTREINA)),;	//35
			Iif(Empty((cTmpAlias)->RA_TELEFON), "NAO INFORMADO", Alltrim((cTmpAlias)->RA_TELEFON)),;	//36
			Iif(Empty((cTmpAlias)->RA_EMAIL),   "NAO INFORMADO", Alltrim((cTmpAlias)->RA_EMAIL)),;	//37
			Iif(Empty((cTmpAlias)->RA_ESTCIVI), "NAO INFORMADO", Alltrim((cTmpAlias)->RA_ESTCIVI)),;	//38
			Iif(Empty((cTmpAlias)->RA_ESTCIVI), "NAO INFORMADO", Alltrim(TABELA("33", (cTmpAlias)->RA_ESTCIVI))),;	//39
			Iif(Empty((cTmpAlias)->RA_NUMCP),   "NAO INFORMADO", Alltrim((cTmpAlias)->RA_NUMCP)),;	//40
			Iif(Empty((cTmpAlias)->RA_DEMISSA), "NAO INFORMADO", Alltrim((cTmpAlias)->RA_DEMISSA)),;	//41
			Iif(Empty((cTmpAlias)->RJ_XAREFUN), "NAO INFORMADO", Alltrim((cTmpAlias)->RJ_XAREFUN)),;	//42
			Iif(Empty((cTmpAlias)->RJ_XDESARE), "NAO INFORMADO", Alltrim((cTmpAlias)->RJ_XDESARE));	//43
			}
			oObjResp  := DadosResponseHBARE120():New(aDados)
			cJsonResp := FWJsonSerialize(oObjResp,.F.,.T.)
			If (cTmpAlias)->CANC == "S"
				lCancel := .T.

				//Verifica se ja enviou a inclusão
				cSelect := " SELECT TOP 1 *"
				cSelect +=   " FROM " + RetSqlName("Z91")
				cSelect +=  " WHERE Z91_FILIAL = '" + FWxFilial("Z91",cFilZ90) + "'"
				cSelect +=    " AND Z91_TPOPER = '" + cTpInt120 + "'"
				cSelect +=    " AND Z91_CHAVE  = '" + (cTmpAlias)->ID + "'"
				cSelect +=    " AND Z91_STATUS = '200'"
				aInclusao := JurSql(cSelect, "*")

				//Carrega inclusão
				If Len(aInclusao) == 0

					// Limpa o log 401 da chave
					U_LimpaZ91(cTpInt120,(cTmpAlias)->ID, DtoS(Date()),'401', .F.,cFilZ90) // .T. Aciona o Like / .F. Busca chave exata
					cSQL := "UPDATE "+RetSqlName("SRA")+" SET RA_MSEXP = '"+DtoS(dDataBase)+"'"
					cSQL += " WHERE RA_FILIAL = '"+FWxFilial("SRA",cFilZ90)+"'"
					cSQL += " AND R_E_C_N_O_  = "+Alltrim(Str((cTmpAlias)->RECSRA))
					TCSQLExec(cSQL)
					(cTmpAlias)->(dbSkip())
					Loop
				EndIf

				//Carrega cancelamento
				cJsonResp := '{'
				cJsonResp += '"id": "' + Iif(Empty((cTmpAlias)->ID), "NAO INFORMADO", (cTmpAlias)->ID) + '",'
				cJsonResp += '"loja": "' + Iif(Empty((cTmpAlias)->LOJA), "NAO INFORMADO", (cTmpAlias)->LOJA) + '"'
				cJsonResp += '}'
				aPath := {cPathCan, cJsonResp}

			Else	//	Inclusão
				aPath := {cPath, cJsonResp}
				lCancel := .F.
			EndIf
			cJsonResp := aPath[2]
			oRestClient:SetPath(aPath[1])
			oRestClient:SetPostParams(FwNoAccent(cJsonResp))
			For nTime := nTimeOut To 0 Step -1
				lIntegr := oRestClient:POST(aHeader)
				If lIntegr
					oJSon := Nil
					cJsonRet := oRestClient:GetResult()
					FWJsonDeserialize(DecodeUtf8(cJsonRet),@oJSon)
					cCodRet := oJSon:RESULT[1]:STATUS
					cMsgRet := oJSon:RESULT[1]:MENSAGEM
					cCodErr := oJSon:RESULT[1]:CODERRO
					cCodAmb := oJSon:RESULT[1]:AMBIENTE

					// Limpa o log 401 da chave
					If lCancel
						U_LimpaZ91(cTpInt120,(cTmpAlias)->ID, DtoS(Date()),'401', .F., cFilZ90) // .T. Aciona o Like / .F. Busca chave exata
						U_LimpaZ91(cTpInt120,(cTmpAlias)->ID, DtoS(Date()),'999', .F., cFilZ90) // .T. Aciona o Like / .F. Busca chave exata
						cTpInt120 := cTpCanc
					Else
						cTpInt120 := GetMV("FS_120TPOP")
						U_LimpaZ91(cTpInt120,(cTmpAlias)->ID, DtoS(Date()),'401', .F., cFilZ90) // .T. Aciona o Like / .F. Busca chave exata
						U_LimpaZ91(cTpInt120,(cTmpAlias)->ID, DtoS(Date()),'999', .F., cFilZ90) // .T. Aciona o Like / .F. Busca chave exata
					EndIf
					Exit

				ELse	//	Erro
					oJSon := Nil
					cJsonRet := oRestClient:GetResult()
					If !Empty(cJsonRet)
						FWJsonDeserialize(DecodeUtf8(cJsonRet), @oJSon)
						If '"status":' $ cJsonRet
							cCodRet := oJSon:RESULT[1]:STATUS
							cMsgRet := oJSon:RESULT[1]:MENSAGEM
							cCodErr := oJSon:RESULT[1]:CODERRO
							cCodAmb := oJSon:RESULT[1]:AMBIENTE
						Else
							cCodRet := '999'
							cMsg := "Retorno inesperado do sistema Arena - "
							If ValType(oJSon:ERROR) == "C"
								cMsg += oJSon:ERROR
							Else
								cMsg += "Retorno ARENA inválido"
							Endif
							cMsgRet := cMsg
							cCodErr := ""
							cCodAmb := ""
						EndIf
					Else
						cCodRet := "999"
						cMsgRet := oRestClient:GetLastError()
						If Empty(cMsgRet)
							If Type("cJsonRet") == "U"
								cMsgRet := "Retorno inesperado do sistema Arena."
							Else
								cMsgRet := Rtrim(cJsonRet)
							EndIf
						EndIf
						cCodErr  := ""
						cCodAmb  := ""
						cJsonRet := "Erro de comunicação"
					EndIf
					If lCancel
						cTpInt120 := cTpCanc
					Else
						cTpInt120 := GetMV("FS_120TPOP")
					EndIf
				EndIf
				If cCodRet == '401'
					Exit
				Endif
				Sleep(1000)	//	Pause de 2 segundos
			Next
			lCorte := (cTmpAlias)->CANC == "S" .And. cCodRet == '401' .And. cDtCorte > (cTmpAlias)->RA_DEMISSA
			If !lCorte
				aEnvZ91 := {}  // Matriz para envio dos campos e valores
				lRet := .T.
				aAdd(aEnvZ91, {"Z91_FILIAL", FWxFilial("Z91",cFilZ90)})
				aAdd(aEnvZ91, {"Z91_TPOPER", cTpInt120})
				aAdd(aEnvZ91, {"Z91_CHAVE" , (cTmpAlias)->ID})
				aAdd(aEnvZ91, {"Z91_JSONOR", cJsonResp})
				aAdd(aEnvZ91, {"Z91_JSONRE", cJsonRet})
				aAdd(aEnvZ91, {"Z91_STATUS", cCodRet})
				aAdd(aEnvZ91, {"Z91_MENSAG", cMsgRet})
				aAdd(aEnvZ91, {"Z91_PROCES", ""})
				If Z91->(FieldPos("Z91_DTORI")) > 0
					aAdd(aEnvZ91, {"Z91_DTORI" , STOD((cTmpAlias)->RA_ADMISSA)})
					aAdd(aEnvZ91, {"Z91_FORCLI", (cTmpAlias)->RA_MAT})
					aAdd(aEnvZ91, {"Z91_TAB"	, "SRA"})
					aAdd(aEnvZ91, {"Z91_INDICE", "1"})
					aAdd(aEnvZ91, {"Z91_CHVP"  , (cTmpAlias)->(RA_FILIAL+RA_MAT)})
					If !Empty(cCodErr)
						aAdd(aEnvZ91, {"Z91_ERRORI", cCodErr})
					EndIf
					If !Empty(cCodAmb)
						aAdd(aEnvZ91, {"Z91_AMB", cCodAmb})
					EndIf
				EndIf
				lRet := U_HBGRV291(aEnvZ91,Nil,@cMsgRet)
			Endif
			If lIntegr .Or. lCorte
				SRA->(DbGoTo((cTmpAlias)->RECSRA))
				SRA->(Reclock('SRA'))
				SRA->RA_MSEXP := DtoS(DATE())
				SRA->(msUnLock())
			EndIf
			Conout("[HBARE020] - " + cCodRet + ": " + (cTmpAlias)->ID + "." + (cTmpAlias)->RA_NOME)
			ASize(aInclusao, 0)
			ASize(aPath, 0)
			FwFreeObj(oObjResp)
			oObjResp := Nil
			(cTmpAlias)->(dbSkip())
		EndDo
	END SEQUENCE
	(cTmpAlias)->(dbCloseArea())

Return


/*/{Protheus.doc} CLASS response
Classes para montar objeto de response
@author Edson Melo
@since 16/04/2020
@type class
/*/
Class DadosResponseHBARE120
	Data id
	Data loja
	Data cpf
	Data rg
	Data nome
	Data apelido
	Data endereco
	Data complemento
	Data bairro
	Data cidade
	Data uf
	Data uf_desc
	Data cep
	Data data_nascimento
	Data setor_id
	Data setor
	Data departamento_id
	Data departamento
	Data turno_id
	Data turno
	Data folga_id
	Data folga
	Data regprof_id
	Data regprof
	Data funcao_id
	Data funcao
	Data numero_registro
	Data data_admissao
	Data sexo_id
	Data sexo
	Data situacao_id
	Data situacao
	Data salario
	Data desc_atividade
	Data treinador
	Data telefone
	Data email
	Data estado_civil_id
	Data estado_civil
	Data ctps
	Data data_demissao
	Data funcao_arena
	Data funcao_arena_desc
	Method New(aDados)

EndClass


Method New(aDados) Class DadosResponseHBARE120
	self:id                := Upper(NoAcento(Alltrim(aDados[01])))
	self:loja              := Upper(NoAcento(Alltrim(aDados[02])))
	self:cpf               := Upper(NoAcento(Alltrim(aDados[03])))
	self:rg                := Upper(NoAcento(Alltrim(aDados[04])))
	self:nome              := Upper(NoAcento(Alltrim(aDados[05])))
	self:apelido           := Upper(NoAcento(Alltrim(aDados[06])))
	self:endereco          := Upper(NoAcento(Alltrim(aDados[07])))
	self:complemento       := Upper(NoAcento(Alltrim(aDados[08])))
	self:bairro            := Upper(NoAcento(Alltrim(aDados[09])))
	self:cidade            := Upper(NoAcento(Alltrim(aDados[10])))
	self:uf                := Upper(NoAcento(Alltrim(aDados[11])))
	self:uf_desc           := Upper(NoAcento(Alltrim(aDados[12])))
	self:cep               := Upper(NoAcento(Alltrim(aDados[13])))
	self:data_nascimento   := Upper(NoAcento(Alltrim(aDados[14])))
	self:setor_id          := Upper(NoAcento(Alltrim(aDados[15])))
	self:setor             := Upper(NoAcento(Alltrim(aDados[16])))
	self:departamento_id   := Upper(NoAcento(Alltrim(aDados[17])))
	self:departamento      := Upper(NoAcento(Alltrim(aDados[18])))
	self:turno_id          := Upper(NoAcento(Alltrim(aDados[19])))
	self:turno             := Upper(NoAcento(Alltrim(aDados[20])))
	self:folga_id          := Upper(NoAcento(Alltrim(aDados[21])))
	self:folga             := Upper(NoAcento(Alltrim(aDados[22])))
	self:regprof_id        := Upper(NoAcento(Alltrim(aDados[23])))
	self:regprof           := Upper(NoAcento(Alltrim(aDados[24])))
	self:funcao_id         := Upper(NoAcento(Alltrim(aDados[25])))
	self:funcao            := Upper(NoAcento(Alltrim(aDados[26])))
	self:numero_registro   := Upper(NoAcento(Alltrim(aDados[27])))
	self:data_admissao     := Upper(NoAcento(Alltrim(aDados[28])))
	self:sexo_id           := Upper(NoAcento(Alltrim(aDados[29])))
	self:sexo              := Upper(NoAcento(Alltrim(aDados[30])))
	self:situacao_id       := Upper(NoAcento(Alltrim(aDados[31])))
	self:situacao          := Upper(NoAcento(Alltrim(aDados[32])))
	self:salario           := aDados[33]
	self:desc_atividade    := Upper(NoAcento(Alltrim(aDados[34])))
	self:treinador         := Upper(NoAcento(Alltrim(aDados[35])))
	self:telefone          := Upper(NoAcento(Alltrim(aDados[36])))
	self:email             := Upper(NoAcento(Alltrim(aDados[37])))
	self:estado_civil_id   := Upper(NoAcento(Alltrim(aDados[38])))
	self:estado_civil      := Upper(NoAcento(Alltrim(aDados[39])))
	self:ctps              := Upper(NoAcento(Alltrim(aDados[40])))
	self:data_demissao     := Upper(NoAcento(Alltrim(aDados[41])))
	self:funcao_arena      := Upper(NoAcento(Alltrim(aDados[42])))
	self:funcao_arena_desc := Upper(NoAcento(Alltrim(aDados[43])))

Return Nil

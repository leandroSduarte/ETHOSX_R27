#INCLUDE "PROTHEUS.CH"
#INCLUDE "TOPCONN.CH"

#DEFINE TPOPERACAO "19"

Static cContArena := ""
Static aIdFat := {}
Static cUrlAre := ""
Static cUrlAglu := ""
Static cUser := ""
Static cPass := ""
Static cID := ""   //	Campo Chave do arquivo Json

/*/ History: 
@P.C.
@	23/12/2020; Aglutinação

@Mário Augusto Cavenaghi
@	Incluído opção para rodar em Console
@	Variável "cID" declarada com STATIC, para inclusão no CONOUT()
@	09/06/21; Utilizar o campo E1_XCHAVE como ID da baixa, e limpeza do código
@	17/06/21; Substituída a função U_GetChvE1(HBAREFUN) por U_GetChSE1(HBARE250), e tratamento do E5_MOTBX = 'DSD' (Desmembramento)
@	02/07/21; Não integrava Fatura aglutinada
/*/

/*/{Protheus.doc} HBAre080()
Job para Envio dos Titulos de Contas a Pagar do financeiro para a integração com Arena
@author Zema
@version P12.1.17
@since 24/09/2020
@return Sem retorno 
/*/
User Function HBARE080(aParam)
	Private cHB_HBARE080 := ""
	Private cHB_NUMRA    := ""
	Private cEX_RCGCMCT  := ""

	Default aParam := {cEmpAnt, cFilAnt}

	If cEmpAnt <> aParam[1]
		If IsBlind()
			RpcSetType(3)
			RpcSetEnv(aParam[1], aParam[2])
		Else
			cEmpAnt := aParam[1]
			cNumEmp := cEmpAnt + cFilAnt
			dbCloseAll()
			OpenSM0(cNumEmp)
			OpenFile(cNumEmp)
		Endif
	Elseif cFilAnt <> aParam[2]
		cFilAnt := aParam[2]
		cNumEmp := cEmpAnt + cFilAnt
	Endif
	cHB_HBARE080 := Alltrim(SuperGetMv("HB_HBARE080", .F., "000210977"))
	cHB_NUMRA    := Alltrim(SuperGetMv("HB_NUMRA"   , .F., "000000248"))
	cEX_RCGCMCT  := Alltrim(SuperGetMv("EX_RCGCMCT" , .F., "22749835"))
	HBARE080Pr("SE2")

RETURN


/*/{Protheus.doc} HBAre081()
Job para Envio dos Titulos de Contas a receber do financeiro para a integração com Arena
@author Zema
@version P12.1.17
@since 24/09/2020
@return Sem retorno 
@obs
/*/
User Function HBARE081(aParam)
	Private cHB_HBARE080 := ""
	Private cHB_NUMRA    := ""
	Private cEX_RCGCMCT  := ""

	Default aParam := {cEmpAnt, cFilAnt}

	If cEmpAnt <> aParam[1]
		If IsBlind()
			RpcSetType(3)
			RpcSetEnv(aParam[1], aParam[2])
		Else
			cEmpAnt := aParam[1]
			cNumEmp := cEmpAnt + cFilAnt
			dbCloseAll()
			OpenSM0(cNumEmp)
			OpenFile(cNumEmp)
		Endif
	Elseif cFilAnt <> aParam[2]
		cFilAnt := aParam[2]
		cNumEmp := cEmpAnt + cFilAnt
	Endif
	Z91->(dbSetOrder(1))
	Z96->(dbSetOrder(1))
	cHB_HBARE080 := Alltrim(SuperGetMv("HB_HBARE080", .F., "000210977"))
	cHB_NUMRA    := Alltrim(SuperGetMv("HB_NUMRA"   , .F., "000000248"))
	cEX_RCGCMCT  := Alltrim(SuperGetMv("EX_RCGCMCT" , .F., "22749835"))
	HBARE080Pr("SE1")

Return Nil


/*/{Protheus.doc} HBARE080Pr
@author Zema
@since 24/09/2020
@type Function
/*/ 
Static Function HBARE080Pr(pAlias)
	Local cTmpAlias := GetNextAlias()
	Local lArenaOn  := GetMV("FS_AREON")
	Local aEmpFil   := {}

	Private cError  := ""
	Private bError  := ErrorBlock({ |oError| cError := oError:Description+CRLF+oError:ErrorStack})
	Private cFilZ90

	cFuncao  := Iif(Alltrim(pAlias)="SE2",PADR("HBARE080",TamSX3("Z97_ROTINA")[1]),PADR("HBARE081",TamSX3("Z97_ROTINA")[1]))
	aStatSem := U_HBVERSEMAF(cFuncao)

	// Verifica se existe registro de semáforo da rotina
	If !aStatSem[1]
		U_HBSMLog("SM",cFilAnt,cFuncao,"499","","","","Codigo chave da rotina "+cFuncao+" nao encontrada na tabela Z97.","499", aStatSem[2] )
		Return
	EndIf

	// Verifica semáforo aberto
	If aStatSem[3] == "2"
		U_HBSMLog("SM",cFilAnt,aStatSem[4],"499","","","","Tentativa de execucao com Semaforo FECHADO","499", aStatSem[2] )
		Return
	EndIf

	// Fecha o semáforo para execução
	aSemaforo := U_HBGRVSEMAF("2", aStatSem[2], aStatSem[4])
	If !aSemaforo[1]
		U_HBSMLog("SM",cFilAnt,aStatSem[4],"499","","","",aSemaforo[2],"499", aStatSem[2] )
		Return
	EndIf

	// Apaga o controle de semáforo fechado
	U_LimpaZ91("SM",aStatSem[2], Dtos(Date()),'499')

	//Seleciona os processos assinados
	If !lArenaOn
		ConOut('[HBARE080Pr] FS_AREON desabilitado.')
		Return
	EndIf

	U_HBJOB01A(cTmpAlias,TPOPERACAO)

	BEGIN SEQUENCE
		While (cTmpAlias)->(!Eof())
			ConOut("[HBARE080Pr] Chamando HBARE801: "+pAlias+" : "+(cTmpAlias)->Z90_EMPRES+"/"+(cTmpAlias)->Z90_FILIAL)
			aEmpFil := {(cTmpAlias)->Z90_EMPRES, (cTmpAlias)->Z90_FILIAL}
			cFilZ90 := (cTmpAlias)->Z90_FILIAL
			U_HBARE801(aEmpFil,pAlias,(cTmpAlias)->Z90_AREFIL)
			(cTmpAlias)->(DbSkip())
		EndDo

		RECOVER
		ConOut("HBARE080 - Erro Fatal")
	END SEQUENCE

	//Restaurando bloco de erro do sistema
	ErrorBlock(bError)

	//Se houve erro de processamento, será gravado na Z91
	If !Empty(cError)
		U_HBSMLog("SM",cFilAnt,"HBARE080","499","","",,"Processamento da rotina HBARE080 na filial "+AllTrim(cFilZ90)+" Interrompido. Vefique o LOG do serviço de JOB Schedule.","499", "0033", .T. )
		ConOut(" "+CRLF+CRLF)
		ConOut("HBARE080- Houve um erro durante a transmissão na FILIAL: "+AllTrim(cFilZ90)+CRLF+cError)
		ConOut(" "+CRLF+CRLF)
	EndIf
	(cTmpAlias)->(dbCloseArea())

	// Abre o semáforo após execução
	aSemaforo := U_HBGRVSEMAF("1", aStatSem[2], aStatSem[4])
	If !aSemaforo[1]
		U_HBSMLog("SM",cFilAnt,aStatSem[4],"499","","","",aSemaforo[2],"499", aStatSem[2] )
	EndIf
	aSize(aStatSem,0)
	aSize(aSemaforo,0)

Return


/*/{Protheus.doc} HBARE801
@author Zema
@since 24/09/2020
@type Function
/*/ 
User Function HBARE801(aEmpFil, pAlias, pLoja)

	Conout("[HBARE801] - inicio do Job de exportacao: "+cEmpAnt+"-"+cFilZ90+" Tabela: "+pAlias)
	HBGerFin(pAlias,pLoja)

Return


/*/{Protheus.doc} HBAre080()
Selecão dos titulos do contas a pagar/ contas a receber para integração com Arena
@param cAlias tabela para geração do json para integração
@author Rene julian
@version P12.1.17
@since 15/04/2020
@return Sem retorno 
@obs
/*/    
Static Function HBGerFin(cAlias,cLoja) //RENE
	Local nQMaxReg := SuperGetMv("FS_AREFIL",,500)  //Quantidade máxima para envio da Integração

	Private cAliasTRB := GetNextAlias()

	If HBSelect(cAlias,cAliasTRB,nQMaxReg)
		HBMontJson(cAlias,cAliasTRB,cLoja)
	Endif
	If Select(cAliasTRB) > 0
		(cAliasTRB)->(DbCloseArea())
	Endif

	//Rotina de envio de cancelamentos das baixas
	HBAreCanBx(cAlias,cLoja)

Return


/*/{Protheus.doc} HBSelect()
Retorna o Alias temporario com os registros selecionados
@param cAlias tabela para geração do json para integração
@param cAliasTRB alias temporario para os registros
@author Rene julian
@version P12.1.17
@since 15/04/2020
@return Sem retorno 
@obs
/*/ 
Static Function HBSelect(cAlias,cAliasTRB,nQMaxReg )
	Local cTopMax  := AllTrim(Str(nQMaxReg))
	Local cCampos1 := ""
	Local cCampos2 := ""
	Local cUniao   := SuperGetMv("MV_UNIAO")
	Local xDataV   := U_HBDTAFECH(TPOPERACAO)

	If Empty(xDataV )
		xDataV := CtoD("01/01/2020")
	EndIf

	If cAlias == "SE2"
		cCampos1 += "% TOP " + cTopMax
		cCampos1 += " SE2.E2_FILIAL,SE2.E2_PREFIXO,SE2.E2_NUM,SE2.E2_FATURA,SE2.E2_PARCELA,SE2.E2_FORNECE,SE2.E2_LOJA,SE2.E2_TIPO,SE2.E2_VALLIQ,SE2.E2_VENCTO,SE2.E2_FILORIG"
		cCampos1 += ",SE2.E2_NATUREZ,SE2.E2_EMISSAO,SE2.E2_EMIS1,SE2.E2_VALOR,SE2.E2_HIST,SE2.E2_BAIXA,SE2.E2_LOTE,SE2.R_E_C_N_O_  REGSE2"
		cCampos1 += ",SE5.E5_BANCO,SE5.E5_AGENCIA,SE5.E5_CONTA,SE5.E5_MOTBX,SE5.E5_VALOR,SE5.E5_AGLIMP,SE5.E5_DATA,SE5.E5_SEQ,SE5.R_E_C_N_O_ REGSE5"
		cCampos1 += ",SE2A.E2_PREFIXO PREFPAI,SE2A.E2_NUM NUMPAI,SE2A.E2_TIPO TIPOPAI,SE2A.E2_PARCELA PARCPAI"
		cCampos2 := cCampos1
		cCampos1 += ",SA2.A2_CGC,SA2.A2_NOME,SA2.A2_NREDUZ %"
		cCampos2 += ",' ' A2_CGC,'"+ cUniao + "' A2_NOME,'"+ cUniao +  "' A2_NREDUZ %"
		BeginSql alias cAliasTRB
         Column E5_DATA AS DATE 
         SELECT %exp:cCampos1%
         FROM   %table:SE2%  SE2             
         INNER  JOIN %table:SE5% SE5
            ON  SE2.E2_FILORIG = SE5.E5_FILORIG 
            AND SE2.E2_PREFIXO = SE5.E5_PREFIXO 
            AND SE2.E2_NUM     = SE5.E5_NUMERO 
            AND SE2.E2_PARCELA = SE5.E5_PARCELA 
            AND SE2.E2_TIPO    = SE5.E5_TIPO 
            AND SE2.E2_FORNECE = SE5.E5_CLIFOR 
            AND SE2.E2_LOJA    = SE5.E5_LOJA 
            AND SE2.E2_BAIXA   = SE5.E5_DATA
         INNER  JOIN %table:SE2%  SE2A 
            ON  SE2A.E2_FILORIG = SE2.E2_FILORIG 
            AND SE2.E2_FATURA  = SE2A.E2_NUM 
            AND SE2.E2_BAIXA   = SE2A.E2_EMISSAO 
            AND SE2A.%notDel%
         INNER  JOIN %table:SA2% SA2
            ON  SA2.A2_FILIAL  = %exp:FWxFilial("SA2",cFilZ90)% 
            AND SA2.A2_COD     = SE2.E2_FORNECE
            AND SE2.E2_LOJA    = SA2.A2_LOJA 
            AND SA2.%notDel%
         WHERE SE2.E2_FILIAL   = %exp:FWxFilial("SE2",cFilZ90)% 
            AND SE2.E2_BAIXA  <> ' ' 
            AND SE5.E5_DATA    > %exp:xDataV%
            AND SE2.E2_FATURA <> ' ' 
            AND SE5.E5_MSEXP   = ' ' 
            AND SE5.E5_XCHAVE  = ' '
            AND SE2.%notDel% 
            AND SE5.%notDel%

		UNION

         SELECT %exp:cCampos2%
         FROM   %table:SE2%  SE2  
         INNER  JOIN %table:SE5% SE5
            ON  SE2.E2_FILORIG  = SE5.E5_FILORIG 
            AND SE2.E2_PREFIXO  = SE5.E5_PREFIXO 
            AND SE2.E2_NUM      = SE5.E5_NUMERO 
            AND SE2.E2_PARCELA  = SE5.E5_PARCELA 
            AND SE2.E2_TIPO     = SE5.E5_TIPO 
            AND SE2.E2_FORNECE  = SE5.E5_CLIFOR 
            AND SE2.E2_LOJA     = SE5.E5_LOJA
         INNER  JOIN %table:SE2% SE2A 
            ON  SE2.E2_FILORIG  = SE2.E2_FILORIG 
            /* AND SE2.E2_AGLIMP   = SE2A.E2_NUM */
				AND RTRIM(SE2.E2_AGLIMP) + RTRIM(SE2.E2_FATURA) = SE2A.E2_NUM
            AND SE2.E2_BAIXA    = SE2A.E2_EMIS1 
            AND SE2.E2_NATUREZ  = SE2A.E2_NATUREZ 
            AND SE2A.%notDel% 
         WHERE  SE2.E2_FILIAL   = %exp:FWxFilial("SE2",cFilZ90)% 
            AND SE2.E2_BAIXA   <> ' ' 
				/*AND SE5.E5_AGLIMP  <> ' '*/
				AND SE2.E2_AGLIMP + SE2.E2_FATURA > ' '
            AND SE5.E5_DATA     > %exp:xDataV%  
            AND SE5.E5_MSEXP    = ' '
            AND SE5.E5_XCHAVE   = ' '
            AND SE5.E5_SITUACA <> 'C' 
            AND SE2.%notDel% 
            AND SE5.%notDel%

         ORDER BY SE2.E2_LOTE, SE5.E5_AGLIMP, SE2.E2_VENCTO, SE2.E2_NUM, SE2.E2_PREFIXO
		EndSql
	Else
		cCampos := '%'+"TOP " + cTopMax
		cCampos += " SE1.E1_FILIAL,SE1.E1_FILORIG,SE1.E1_PREFIXO,SE1.E1_NUM,SE1.E1_FATURA,SE1.E1_PARCELA,SE1.E1_NATUREZ,SE1.E1_CLIENTE,SE1.E1_LOJA"
		cCampos += ",SE1.E1_TIPO,SE1.E1_VALLIQ,SE5.E5_BANCO,SE5.E5_AGENCIA,SE5.E5_CONTA,SE1.E1_LOTE,SE1.R_E_C_N_O_ REGSE1,SE5.E5_MOTBX,SE5.E5_AGLIMP"
		cCampos += ",SE5.E5_DOCUMEN, SA1.A1_CGC,SA1.A1_NOME,SA1.A1_NREDUZ,SE1.E1_EMISSAO,SE1.E1_EMIS1,SE1.E1_VALOR,SE1.E1_HIST,SE5.E5_VALOR,SE1.E1_BAIXA"
		cCampos += ",SE1.E1_VENCTO,SE5.E5_DATA,SE5.R_E_C_N_O_ REGSE5,SE5.E5_SEQ,E1_IRRF, E1_CSLL, E1_PIS, E1_COFINS, E1_XCHAVE " + '%
		BeginSql alias cAliasTRB
         Column E1_BAIXA AS DATE 
         Column E5_DATA AS DATE 
         SELECT %exp:cCampos%
         FROM   %table:SE5% SE5
         INNER  JOIN %table:SE1% SE1
            ON  SE1.E1_FILORIG = SE5.E5_FILORIG 
            AND SE1.E1_PREFIXO = SE5.E5_PREFIXO 
            AND SE1.E1_NUM     = SE5.E5_NUMERO 
            AND SE1.E1_TIPO    = SE5.E5_TIPO 
            AND SE1.E1_CLIENTE = SE5.E5_CLIFOR 
            AND SE1.E1_PARCELA = SE5.E5_PARCELA
				AND SE1.%notDel% 
         INNER  JOIN %table:SA1% SA1
            ON  A1_FILIAL      = %exp:FWxFilial("SA1",cFilZ90)% 
            AND SA1.A1_COD     = SE1.E1_CLIENTE 
            AND SE1.E1_LOJA    = SA1.A1_LOJA 
            AND SA1.%notDel%
         WHERE  SE1.E1_FILIAL  = %exp:FWxFilial("SE2",cFilZ90)% 
            AND SE5.E5_RECPAG  = 'R' 
            AND SE5.E5_DATA    > %exp:xDataV% 
            AND SE5.E5_TIPODOC = 'BA' 
            AND SE5.E5_MSEXP   = ' '
            AND SE5.E5_XCHAVE  = ' '            
            AND SE5.E5_SITUACA <> 'C'
            //	Aglutinação
            AND ((SE5.E5_MOTBX IN ('LIQ', 'DSD') AND SE5.E5_TIPODOC = 'BA') 
            OR  (E5_FILORIG = SE1.E1_FILORIG AND E5_ORIGEM ='GSFINFT1' AND E5_PREFIXO =  '0' AND E5_TIPO = 'NF' AND E1_NUMRA <> ' ' AND E5_MOTBX = 'MCT' AND E5_TIPODOC = 'BA' AND LEFT(A1_CGC,8) = %exp:cEX_RCGCMCT%)
            OR  (E5_FILORIG = SE1.E1_FILORIG AND E5_ORIGEM ='U_AGLTIT' AND E5_TIPO = 'NF' AND E5_PREFIXO IN ('RPS','ND') AND E5_MOTBX IN('LIQ','NOR') AND E5_TIPODOC IN('VL','BA')))         
            //=====================
            AND SE5.%notDel%  
				//AND SE5.E5_NUMERO = '01042021A'	//	REMOVER
         ORDER BY E5_DATA,E5_PREFIXO,E5_NUMERO,E5_PARCELA,E5_CLIENTE,E5_LOJA
		EndSql
	Endif
	aQuery := GetLastQuery()

Return(!(cAliasTRB)->(Eof()))


/*/{Protheus.doc} HBIDFat()  
Retorna o dados para a geração do ID do contas a receber
@param cNumLiq Codigo que foi realizado a Liquidez 
@param dDtEmiss data do processamento da Liquidez
@author Rene julian
@version P12.1.17
@since 12/06/2020
@return aRet array com as informações do registros de aglutinação
@obs
/*/    
Static Function HBIDFat(cNumLiq,DtEmiss)
	Local aRet       := {}              //Retorno do codigo do registro de aglutinação
	Local cAliasTRB3 := GetNextAlias()  //Alias da query

	If Select(cAliasTRB3) > 0
		(cAliasTRB3)->(DbCloseArea())
	EndIf

	BeginSql alias cAliasTRB3
      SELECT E1_PREFIXO, E1_NUM, E1_PARCELA, E1_TIPO 
      FROM   %table:SE1% SE1
      WHERE  SE1.E1_FILORIG = %exp:cFilZ90% 
         AND SE1.E1_NUMLIQ  = %exp:cNumLiq% 
         AND SE1.E1_EMISSAO = %exp:DtEmiss% 
         AND SE1.%notDel%
	EndSql

	If !(cAliasTRB3)->(Eof())
		aAdd(aRet, (cAliasTRB3)->E1_PREFIXO )
		aAdd(aRet, (cAliasTRB3)->E1_NUM )
		aAdd(aRet, (cAliasTRB3)->E1_PARCELA )
		aAdd(aRet, (cAliasTRB3)->E1_TIPO )
	EndIf

Return aRet


/*/{Protheus.doc} HBMontJson()
Monto o arquivo no formato Json para a integração com Arena
@param cAlias tabela para geração do json para integração
@param cAliasTRB Alias com os registros a serem enviados para a integração
@author Rene julian
@version P12.1.17
@since 16/04/2020
@return 
@obs
/*/
Static Function HBMontJson(cAlias,cAliasTRB,cLoja)
	Local cJSon     := ""   //Retorno com o arquivo para envio.
	Local cMsgJSon  := ""   // Mensagem de Retorno do JSon
	Local aAtuAlias := {}   // Array para atualização do campo MSEXP
	Local lRetJSon  := .T.  // Retorno do envio das informações JSon
	Local nX        := 0

	While !(cAliasTRB)->(Eof())
		cJSon    := HBCabJson(cAlias,cAliasTRB,@aAtuAlias, cLoja)
		lRetJSon := HBenvJSon(cJSon,@cMsgJSon)
		If lRetJSon
			SE5->(DbGoTo((cAliasTRB)->REGSE5))
			RecLock("SE5", .F.)
			SE5->E5_MSEXP  := DtoS(Date())
			SE5->E5_XCHAVE := aAtuAlias[1]
			SE5->(MsUnlock())
			aCodRet := HBRetMsg(cMsgJSon) //ok
		Else
			If ValType(cMsgJSon) == "U"
				Conout("Erro cMsgJSon")
			EndIf
			aCodRet := HBRetMsg(cMsgJSon) //erro
		EndIf
		HBGrvZ91(cJSon, aCodRet, cMsgJSon, aAtuAlias[1], cAlias,, cAliasTRB)
		ConOut("[HBARE080] " + cEmpAnt + "/" + cFilAnt + " " + cID)
		aAtuAlias := {}
		(cAliasTRB)->(DbSkip())
	Enddo

Return Nil


/*/{Protheus.doc} HBIntArena()
Retorno das informações de tabelas De para
@param cAlias tabela para geração do json para integração
@param cAliasTRB Alias com os registros a serem enviados para a integração
@author Rene julian
@version P12.1.17
@since 16/04/2020
@return 
@obs
/*/ 
Static Function HBIntArena(cAlias,cFilOrig,cMotBx,cAGLImp)
	Local aRet      := {} //Retorno com dos campos De para com Arena
	Local ctipo     := Iif(cAlias == "SE2","19","20")  //Vai ser usando o cod da nomenclarua do fonte. - Tipo de integração (tabela inserida no SX5)
	Local cRecPag   := Iif(cAlias == "SE2","AP","AR")
	Local cConArena := ""

	//Dados de Agencia e banco Arena
	If !Empty(cAGLImp) .And. Empty(cContArena)
		cContArena := SuperGetMv("MV_ARENACTA",,"")
	EndIf
	If !Empty(cAGLImp) .And. !Empty(cContArena)
		cConArena := cContArena
	Else
		Z93->(dbSetOrder(2))
		Z93->(MsSeek(xFilial("Z93",cFilZ90)+cMotBx))
		While !Z93->(Eof()) .AND. Upper(Z93->Z93_MOTBX) == cMotBx
			If  Z93->Z93_RECPAG $ cRecPag
				cConArena := AllTrim(Z93->Z93_CONTA)
				EXIT
			EndIf
			Z93->(DbSkip())
		End
	EndIf
	aRet := {ctipo,cConArena}

Return aRet


/*/{Protheus.doc} HBCabJson()
Retorno das informações de cabeçalho para montagen do Json
@param cAlias tabela para geração do json para integração
@param cAliasTRB Alias com os registros a serem enviados para a integração
@author Rene julian
@version P12.1.17
@since 16/04/2020
@return 
@obs
/*/ 
Static Function HBCabJson(cAlias,cAliasTRB,aAtuAlias,cLoja)
	Local aCampJson := {}   //Campos do cabeçalho
	Local aInfoJson := {}   //Informações do cabeçalho
	Local cIDTit    := ""   //Campo Chave do arquivo Json
	Local cIDFat    := ""   //Campo Chave do arquivo Json
	Local aArenaInt := {}   //Recebe os campos de para da integação Arena
	Local cEstorno  := "N"  //Valor fixo para envio dos registro de baixa
	Local oObjJSON  := JsonObject():New()
	Local cEnvJson  := ""   //Retorno com as informações do Cabeçalho JSon
	Local nX        := 0    //Auxiliar para o For
	Local cNumFat   := ""
	Local cPrefFat  := ""
	Local cParcFat  := ""
	Local cTipoFat  := ""
	Local aLiqInf   := ""   //Informações do Titulo de Aglutinação

	Conout("[HBARE801] - Monta Json para envio")

	//Campos
	aCampJson := {"id","idfatura","loja","tipo","conta","cnpjfor","nomefor","nomefan","numcheque","valor","val_orig","estorno"}

	If cAlias == "SE2"
		cID    := (cAliasTRB)->E2_FILORIG + (cAliasTRB)->E2_PREFIXO + (cAliasTRB)->E2_NUM + (cAliasTRB)->E2_PARCELA + (cAliasTRB)->E2_FORNECE + (cAliasTRB)->E2_LOJA + (cAliasTRB)->E2_TIPO
		cIDTit := (cAliasTRB)->E2_FILORIG + (cAliasTRB)->E2_PREFIXO + (cAliasTRB)->E2_NUM + (cAliasTRB)->E2_PARCELA + (cAliasTRB)->E2_FORNECE + (cAliasTRB)->E2_LOJA
		cIDFat := (cAliasTRB)->E2_FILORIG + (cAliasTRB)->PREFPAI + (cAliasTRB)->NUMPAI  + (cAliasTRB)->PARCPAI  + (cAliasTRB)->E2_FORNECE + (cAliasTRB)->E2_LOJA + (cAliasTRB)->TIPOPAI
		aArenaInt := HBIntArena(cAlias,(cAliasTRB)->E2_FILORIG,(cAliasTRB)->E5_MOTBX,(cAliasTRB)->E5_AGLIMP )

		//Informações
		aInfoJson  :=  {cID+'19' ,;
			cIDFat  ,;
			cLoja ,;
			aArenaInt[1]  ,;
			aArenaInt[2]  ,;
			Alltrim((cAliasTRB)->A2_CGC)    ,;
			AllTrim((cAliasTRB)->A2_NOME)   ,;
			Alltrim((cAliasTRB)->A2_NREDUZ) ,;
			"",;
			AllTrim(Transform((cAliasTRB)->E2_VALLIQ,"@R 999999999999.99")) ,;
			AllTrim(Transform((cAliasTRB)->E2_VALOR, "@R 999999999999.99")) ,;
			cEstorno}

		aAdd(aAtuAlias,cID+'19' )
		aAdd(aAtuAlias,(cAliasTRB)->REGSE5)
	Else
		If Len(aIdFat) >= 5 .And. AllTrim(aIdFat[5]) == AllTrim((cAliasTRB)->E5_DOCUMEN)
			cPrefFat := aIdFat[1]
			cNumFat  := aIdFat[2]
			cParcFat := aIdFat[3]
			cTipoFat := aIdFat[4]
		Else
			aLiqInf := HBIDFat((cAliasTRB)->E5_DOCUMEN,(cAliasTRB)->E5_DATA)

			If Len(aLiqInf) >= 4
				aIdFat := {}
				cPrefFat := aLiqInf[1]
				cNumFat  := aLiqInf[2]
				cParcFat := aLiqInf[3]
				cTipoFat := aLiqInf[4]
				aAdd(aIdFat, cPrefFat)
				aAdd(aIdFat, cNumFat )
				aAdd(aIdFat, cParcFat)
				aAdd(aIdFat, cTipoFat)
				aAdd(aIdFat,(cAliasTRB)->E5_DOCUMEN)
			EndIf
		EndIf

		If !Empty((cAliasTRB)->E1_XCHAVE)
			cID := Alltrim((cAliasTRB)->E1_XCHAVE)
		Else
			//cID := U_GetChvE1(cAliasTRB, "10", (cAliasTRB)->REGSE1, cLoja, .F. ) //	+ (cAliasTRB)->E1_PARCELA
			cID := U_GetChSE1(cAliasTRB, "10", (cAliasTRB)->REGSE1, cLoja, .F. )
		Endif
		cIDTit := cID
		cIDFat := (cAliasTRB)->E1_FILORIG + DTOS((cAliasTRB)->E5_DATA) + cPrefFat + cNumFat + cTipoFat + cParcFat
		aArenaInt := HBIntArena(cAlias,(cAliasTRB)->E1_FILORIG,(cAliasTRB)->E5_MOTBX,(cAliasTRB)->E5_AGLIMP )

		//Informações
		aInfoJson  :=  {cID+'20' ,;
			cIDFat  ,;
			cLoja   ,;
			aArenaInt[1]  ,;
			aArenaInt[2] ,;
			Alltrim((cAliasTRB)->A1_CGC)    ,;
			AllTrim((cAliasTRB)->A1_NOME)   ,;
			Alltrim((cAliasTRB)->A1_NREDUZ) ,;
			"",;
			AllTrim(Transform((cAliasTRB)->E1_VALLIQ,"@R 99999999999.99"))  ,;
			AllTrim(Transform((cAliasTRB)->E1_VALOR, "@R 99999999999.99"))  ,;
			cEstorno}

		aAdd(aAtuAlias,cID+'20')
		aAdd(aAtuAlias,(cAliasTRB)->REGSE5)
	EndIf

	For nX := 1 to Len( aCampJson )
		oObjJSON[aCampJson[nX]] := aInfoJson[nX]
	Next
	FWJsonSerialize(oObjJSON)
	cEnvJson := HBTitJson(cAlias,cAliasTRB,oObjJSON,cID,cIDFat, cIDTit)
	FreeObj(oObjJSON)

Return cEnvJson


/*/{Protheus.doc} HBTitJSon()
Retorno das informações de cabeçalho para montagen do Json
@param cAlias tabela para geração do json para integração
@param cAliasTRB Alias com os registros a serem enviados para a integração
@author Rene julian
@version P12.1.17
@since 16/04/2020
@return 
@obs
/*/ 
Static Function HBTitJSon(cAlias,cAliasTRB,oObjJSON,cID,cIDFat,cIDTit)
	Local aCampJson := {}   //Campos do cabeçalho
	Local aInfoJson := {}   //Informações do cabeçalho
	Local cEnvJson  := ""
	Local nImpostos := 0
	Local nX := 0    //Auxiliar para o For
	Local ny := 0    //Auxiliar para o For

	Conout("JOBFIN04 - Informações do titulo do JSon")

	oObjJSON["titulo"] := {}

	If cAlias == "SE2"
		aCampJson := {"valorpago","idfatura","dtbaixa", "parcela", "seq", "idtit"}
		SE2->(DBGOTO((cAliasTRB)->REGSE2))
		aInfoJson := {AllTrim(Transform(SE2->E2_VALLIQ-SE2->E2_PIS-SE2->E2_COFINS-SE2->E2_CSLL,"@R 999999999999.99")), ;
			cIDFat, ;
			DTOS((cAliasTRB)->E5_DATA), ;
			SE2->E2_PARCELA , ;
			(cAliasTRB)->E5_SEQ, ;
			cID }
		aAdd(oObjJSON["titulo"], JsonObject():New() )
		ny++
		For nX := 1 to Len( aCampJson )
			oObjJSON["titulo", ny, aCampJson[nX]] := aInfoJson[nX]
		Next

		If SE2->E2_PIS > 0
			aInfoJson := {AllTrim(Transform(SE2->E2_PIS,"@R 999999999999.99")), ;
				cIDFat, ;
				DTOS((cAliasTRB)->E5_DATA), ;
				SE2->E2_PARCELA , ;
				(cAliasTRB)->E5_SEQ, ;
				cIDTit+'PIS' }

			aAdd(oObjJSON["titulo"], JsonObject():New() )
			ny++
			For nX := 1 to Len( aCampJson )
				oObjJSON["titulo", ny, aCampJson[nX]] := aInfoJson[nX]
			Next
		Endif

		If SE2->E2_COFINS > 0
			aInfoJson := {AllTrim(Transform(SE2->E2_COFINS,"@R 999999999999.99")), ;
				cIDFat, ;
				DTOS((cAliasTRB)->E5_DATA), ;
				SE2->E2_PARCELA , ;
				(cAliasTRB)->E5_SEQ, ;
				cIDTit+'COF' }

			aAdd(oObjJSON["titulo"], JsonObject():New() )
			ny++
			For nX := 1 to Len( aCampJson )
				oObjJSON["titulo", ny, aCampJson[nX]] := aInfoJson[nX]
			Next
		Endif

		If SE2->E2_CSLL > 0
			aInfoJson := {AllTrim(Transform(SE2->E2_CSLL,"@R 999999999999.99")), ;
				cIDFat, ;
				DTOS((cAliasTRB)->E5_DATA), ;
				SE2->E2_PARCELA , ;
				(cAliasTRB)->E5_SEQ, ;
				cIDTit+'CSL' }

			aAdd(oObjJSON["titulo"], JsonObject():New() )
			ny++
			For nX := 1 to Len( aCampJson )
				oObjJSON["titulo", ny, aCampJson[nX]] := aInfoJson[nX]
			Next
		Endif
	Else
		aCampJson := {"idtit","idfatura", "valorrecebido","dtbaixa","parcela","seq","impostos"}
		nImpostos := 0
		If (cAliasTRB)->E1_VALLIQ <> (cAliasTRB)->E1_VALOR
			If Round((cAliasTRB)->(E1_VALLIQ+E1_CSLL+E1_PIS+E1_COFINS),2) = Round((cAliasTRB)->E1_VALOR,2)
				nImpostos := (cAliasTRB)->(E1_CSLL+E1_PIS+E1_COFINS)
			Endif
			If Round((cAliasTRB)->(E1_VALLIQ+E1_IRRF+E1_CSLL+E1_PIS+E1_COFINS),2) = Round((cAliasTRB)->E1_VALOR,2)
				nImpostos := (cAliasTRB)->(E1_IRRF+E1_CSLL+E1_PIS+E1_COFINS)
			Endif
		Endif
		aInfoJson := {cIDTit, ;
			cIDFat, ;
			AllTrim(Transform((cAliasTRB)->E5_VALOR,"@R 999999999999.99")), ;
			DtoS((cAliasTRB)->E1_BAIXA) , ;
			Alltrim((cAliasTRB)->E1_PARCELA) , ;
			Alltrim((cAliasTRB)->E5_SEQ) ,;
			Alltrim(cValToChar(nImpostos))}
		aAdd(oObjJSON["titulo"], JsonObject():New())
		For nX := 1 to Len( aCampJson )
			oObjJSON["titulo", 1, aCampJson[nX]] := aInfoJson[nX]
		Next
	Endif
	cEnvJson := FWJsonSerialize(oObjJSON)

return cEnvJson


/*/{Protheus.doc} HBenvJSon()
Envio das informações via JSon
@param cJSon texto montado para envio Arena
@author Rene julian
@version P12.1.17
@since 20/04/2020
@return 
@obs
/*/ 
Static Function HBenvJSon(cJSon, cMsgJSon, lCancel)
	Local lRet     := .T.   //Retorno da Função indicando se houve erro no envio do arquivo JSon
	Local oRest    := Nil
	Local oRetorno := Nil
	Local cRet     := ""  //retorno do Rest Arena
	Local cErrJson := ""

	Default lCancel := .F.

	Private aHeader := {}

	If Empty(cUrlAre)
		cUrlAre := SuperGetMv("FS_URLMOVB",,"http://ddnss-vmapp.ddns.com.br:7001")
		cUser   := SuperGetMv("FS_USRAREN",,"Protheus")
		cPass   := SuperGetMv("FS_PASWARE",,"!@arena@!")
	EndIf
	If lCancel
		cUrlAglu := SuperGetMv("FS_URLCFIN",,"/Arena/API_01/TSM/AgluTitCanc")
	Else
		cUrlAglu := SuperGetMv("FS_URLBFIN",,"/Arena/API_01/TSM/AgluTit")
	EndIf
	oRest := FwRest():New(cUrlAre)
	aAdd(aHeader, "Content-Type:application/json; Charset=UTF-8" )
	aAdd(aHeader, "Authorization: Basic " + Encode64(cUser + ":" + cPass) )
	oRest:SetPath(cUrlAglu)  //Complemento do endereço da Api Rest
	cJson := FwNoAccent(cJson)
	oRest:SetPostParams(cJSon)
	If oRest:Post(aHeader)
		cMsgJSon := oRest:GetResult()
		If oRetorno == Nil
			oRetorno := JsonObject():New()
			oRetorno:FromJson(cRet)
			FreeObj(oRetorno)
		EndIf
		Conout("[HBARE801] - Envio com Sucesso Mensagem:"+cMsgJSon )
	Else
		lRet := .F.
		cErrJson := oRest:GetResult()
		cMsgJSon := Iif(ValType(cErrJson) == "C" .And. !Empty(cErrJson), cErrJson , oRest:GetlastError() )
		Conout("[HBARE801] - Envio com erro Mensagem:"+cMsgJSon )
	EndIf
	FreeObj(oRest)

Return lRet


/*/{Protheus.doc} HBRetMsg()
Função para tratamento de mensagem de retorno do Json de Erro
@Param cMsgJSon Retorno do Json a ser trnasformado em array
@param cTipo 1 = mensagem de ok | 2 = mensagem de erro 
@author Rene julian
@version P12.1.17
@since 24/05/2020
@return 
@obs
/*/
Static Function HBRetMsg(cMsgJSon)
	Local aRet     := {} //Retorno da Mensagem e status do Arena
	Local oObjJSON := Nil
	Local cJSon    := ""
	Local JSonRST  := Nil

	If !Empty(cMsgJSon)
		oObjJSON   := JsonObject():New()
		cJSon := oObjJSON:FromJson(cMsgJSon)
		If Empty(cJSon)
			JSonRST := oObjJSON:GetJsonObject("result")
			If Valtype(JSonRST) <> "U" .And. Valtype(JSonRST[1]) == "J"
				aAdd(aRet,{'mensagem' , JSonRST[1, "mensagem"] } )
				aAdd(aRet,{'status'   , JSonRST[1, "status"]   } )
				If At("descricao",cMsgJSon) > 0
					aAdd(aRet,{'erro_descr', JSonRST[1, "erros", 1, "descricao"]  } )
					aAdd(aRet,{'erro_id'   , JSonRST[1, "erros", 1, "bloco"]  } )
				Else
					aAdd(aRet,{'erro_descr', ""  } )
					aAdd(aRet,{'erro_id'   , ""  } )
				EndIf
				aAdd(aRet,{'CodErro' , JSonRST[1, "codErro"]   } )
				aAdd(aRet,{'Ambiente', JSonRST[1, "ambiente"]   } )
			Else
				aAdd(aRet,{'mensagem', cMsgJSon } )
			EndIf
		Else
			aAdd(aRet,{'mensagem', cMsgJSon } )
		EndIf
		FreeObj(oObjJSON)
	EndIf

Return aRet


/*/{Protheus.doc} HBGrvZ91()
Esta funcao realiza a gravação da tabela Z91 de logs para 
a integracao com o Arena 
@author Rene Julian
@version P12.1.17
@since 11/05/2020
@return 
@obs 
/*/ 
Static Function HBGrvZ91(cDadJSon,aCodRet,cCodRet,cIDArena,cAlias,cDeleted,cAliasTMP)
	Local aEnvZ91  := {}  // Matriz para envio dos campos e valores
	Local cMsgErro := "" // Retorno da Mensagem de erro de gravação da Z91
	Local lRet     := .T.
	Local cMensag  := ""
	Local cTipOper := Iif(cAlias == "SE2","19","20" )

	Default cDeleted := ""

	If !Empty(cDeleted)
		If cTipOper != cDeleted
			cTipOper := cDeleted
		EndIf
	EndIf
	If Len(aCodRet) > 0
		cMensag := aCodRet[1, 1]+":"+ aCodRet[1, 2]
		If Len(aCodRet) > 2
			cMensag += CRLF + "Descrição:" + aCodRet[3, 2]
		EndIf
	EndIf
	aAdd(aEnvZ91, {"Z91_FILIAL", FWxFilial("Z91",cFilZ90)} )
	aAdd(aEnvZ91, {"Z91_TPOPER", cTipOper})
	aAdd(aEnvZ91, {"Z91_CHAVE" , cIDArena })
	aAdd(aEnvZ91, {"Z91_JSONOR", cDadJSon })
	aAdd(aEnvZ91, {"Z91_DTOPER", Date() })
	aAdd(aEnvZ91, {"Z91_JSONRE", cCodRet })
	If Len(aCodRet) > 1
		aAdd(aEnvZ91, {"Z91_STATUS", aCodRet[2, 2] })
	EndIf
	aAdd(aEnvZ91, {"Z91_MENSAG", cMensag })
	aAdd(aEnvZ91, {"Z91_PROCES", "1" })

	// 10/03/21 - Atilio - Campos novos do monitor
	If Z91->( FieldPos("Z91_DTORI") ) > 0
      /*
      Campos Criados :
      Z91_DTORI  D 8    Data Origem Lançamento
      Z91_ERRORI C 100  Origem do Erro
      Z91_FORCLI C 20   Fornecedor ou Cliente ou Matrícula
      Z91_TITULO C 20   Nro. Do Título
      Z91_VALOR  N 15,2 Valor do Título
      Z91_AMB    C 3    Ambiente (recebe do arena)
      Z91_EMAIL  C 1    ENVIA E-MAIL (monitor)
      */
		SE5->(dbGoTo((cAliasTMP)->REGSE5))
		aAdd(aEnvZ91, {"Z91_DTORI" , SE5->E5_DATA})
		//aAdd(aEnvZ91, {"Z91_ERRORI" , "" 	       })
		aAdd(aEnvZ91, {"Z91_FORCLI", SE5->E5_CLIFOR})
		aAdd(aEnvZ91, {"Z91_TITULO", SE5->E5_NUMERO})
		aAdd(aEnvZ91, {"Z91_VALOR" , SE5->E5_VALOR})
		aAdd(aEnvZ91, {"Z91_TAB"   , "SE5"})
		aAdd(aEnvZ91, {"Z91_INDICE", "L"})
		aAdd(aEnvZ91, {"Z91_CHVP"  , SE5->( E5_FILIAL+E5_IDORIG+E5_TIPODOC )})
		If Len(aCodRet) > 4
			aAdd(aEnvZ91, {"Z91_ERRORI", aCodRet[5, 2]})
		EndIf
		If Len(aCodRet) > 5
			aAdd(aEnvZ91, {"Z91_AMB"   , aCodRet[6, 2]})
		EndIf
	EndIf
	cFilBkp := cFilAnt
	cFilAnt := cFilZ90
	lRet := U_HBGRV291(aEnvZ91,Nil,@cMsgErro)
	cFilAnt := cFilBkp
	If !lRet .And. !Empty(cMsgErro)
		Conout("[HBARE801] - Erro na gravação da Tabela de Log mensagem:"+cMsgErro )
	EndIf

Return


/*/{Protheus.doc} HBAreCanBx()
Função vai buscar os registros para serem posicionados
@author Rene julian
@version P12.1.17
@since 24/06/2020
@return 
@obs
/*/
Static Function HBAreCanBx(cAlias,cLoja)
	Local lRetJSon := .F.
	Local aCodRet  := {}
	Local cEnvJson := ""   //Retorno com as informações do Cabeçalho JSon
	Local cMsgJSon := ""
	Local cID      := ""
	Local cTpOper  := ""
	Local cTpExcl  := ""

	Private cAliasTRB4 := GetNextAlias() //Alias da query

	(cAlias)->(dbSetOrder(1))
	HBSelecCan(cAlias,cAliasTRB4)

	While !(cAliasTRB4)->( Eof() )

		If cAlias == "SE2"
			cID     := (cAliasTRB4)->(E5_FILORIG + E5_PREFIXO + E5_NUMERO + E5_PARCELA + E5_CLIFOR + E5_LOJA + E5_TIPO)+'19'
			cTpOper := "19"
			cTpExcl := GetAdvFVal("Z96","Z96_CHAVEC",FWxFilial("Z96",cFilZ90))
		Else
			//cID := U_GetChvE1(cAliasTRB4, "10", (cAliasTRB4)->REGSE1, cLoja, .F. ) + (cAliasTRB4)->E5_PARCELA +'20'
			cID     := U_GetChSE1(cAliasTRB4, "10", (cAliasTRB4)->REGSE1, cLoja, .F. ) + (cAliasTRB4)->E5_PARCELA +'20'
			cTpOper := "20"
		EndIf
		cTpExcl := Stuff( Space(TamSX3("Z96_CHAVE")[1]) , 1 , Len(AllTrim(cTpOper)) , AllTrim(cTpOper) )
		cTpExcl := AllTrim(Posicione("Z96",1,FWxFilial("Z96")+cTpExcl, "Z96_CHAVEC"))

		// Verifica se a inclusão foi processada no ARENA
		If U_fLerZ91(cFilZ90 ,cID, cTpOper,"401")
			U_LimpaZ91(cTpOper,cID, Dtos(Date()),'401', .F.,cFilZ90) // .T. Aciona o Like / .F. Busca chave exata
			U_LimpaZ91(cTpExcl,cID, Dtos(Date()),'401', .F.,cFilZ90) // .T. Aciona o Like / .F. Busca chave exata
			cSQL := "UPDATE "+RetSqlName("SE5")+" "
			cSQL += "SET E5_MSEXP = '"+DTOS(dDataBase)+"', "
			cSQL += " E5_XCHAVE = '"+cID+"' "
			cSQL += " WHERE E5_FILIAL = '"+xFilial("SE5",cFilZ90)+"'"
			cSQL += " AND R_E_C_N_O_ = "+Alltrim(Str((cAliasTRB4)->REGSE5))
			TCSQLExec(cSQL)
			(cAliasTRB4)->(dbSkip())
			Loop
		EndIf

		lRetJSon := HBAre082(cAlias,@cEnvJson,@cMsgJSon, cID, AllTrim(cLoja))
		If lRetJSon
			U_LimpaZ91(cTpExcl,cID, Dtos(Date()),'401', .F.,cFilZ90) // .T. Aciona o Like / .F. Busca chave exata
			cSQL := "UPDATE "+RetSqlName("SE5")+" "
			cSQL += "SET E5_MSEXP = '"+DTOS(dDataBase)+"', "
			cSQL += " E5_XCHAVE = '"+cID+"' "
			cSQL += " WHERE E5_FILIAL = '"+xFilial("SE5",cFilZ90)+"'"
			cSQL += " AND R_E_C_N_O_ = "+Alltrim(Str((cAliasTRB4)->REGSE5))
			//ConOut(cSQL)
			TCSQLExec(cSQL)
			aCodRet := HBRetMsg(cMsgJSon) //ok
		Else
			if ValType(cMsgJSon) == "U"
				Conout("Erro cMsgJSon")
				cMsgJSon := "Erro no Json Valor Nulo"
			EndIf
			aCodRet := HBRetMsg(cMsgJSon) //erro
		EndIf
		HBGrvZ91(cEnvJson,aCodRet,cMsgJSon,cID,cAlias,cTpExcl,cAliasTRB4)
		(cAliasTRB4)->(DbSkip())
	Enddo
	(cAliasTRB4)->(DbCloseArea())

Return


/*/{Protheus.doc} HBSelecCan()  
Retorna o dados para a geração do ID do contas a receber
@param cNumLiq Codigo que foi realizado a Liquidez 
@param dDtEmiss data do processamento da Liquidez
@author Rene julian
@version P12.1.17
@since 12/06/2020
@return aRet array com as informações do registros de aglutinação
@obs
/*/  
Static Function HBSelecCan(cAlias,cAliasTRB4)
	Local xDataV := U_HBDTAFECH(TPOPERACAO)

	If cAlias == "SE2"
		BeginSql alias cAliasTRB4
         SELECT E5_FILIAL,E5_FILORIG,E5_PREFIXO,E5_NUMERO,E5_PARCELA, E5_CLIFOR,E5_LOJA,E5_TIPO, SE5.E5_DATA, SE5.R_E_C_N_O_  REGSE5
         FROM   %table:SE5% SE5
         WHERE  SE5.E5_FILORIG = %exp:cFilZ90% 
            AND SE5.E5_TIPO    = 'TX' 
            AND SE5.E5_AGLIMP <> ' ' 
            AND SE5.E5_SITUACA = 'C' 
            AND SE5.E5_MSEXP   = ' ' 
            AND SE5.E5_DATA    > %exp:xDataV%

      UNION

         SELECT E5_FILIAL,E5_FILORIG,E5_PREFIXO,E5_NUMERO,E5_PARCELA, E5_CLIFOR,E5_LOJA,E5_TIPO,SE5.E5_DATA,  SE5.R_E_C_N_O_  REGSE5
         FROM   %table:SE5% SE5 
         WHERE  SE5.E5_FILORIG = %exp:cFilZ90%  
            AND SE5.D_E_L_E_T_ = '*'
            AND SE5.E5_MOTBX   = 'FAT' 
            AND SE5.E5_MSEXP   = ' '
            AND SE5.E5_SITUACA = 'C'  
            AND SE5.E5_DATA    > %exp:xDataV%
         ORDER  BY E5_TIPO,E5_PREFIXO,E5_NUMERO,E5_PARCELA, E5_CLIFOR,E5_LOJA
		EndSql

	Else
		BeginSql alias cAliasTRB4
         SELECT E5_FILIAL, E5_FILORIG, E5_PREFIXO, E5_NUMERO, E5_PARCELA, E5_CLIFOR, E5_LOJA, E5_TIPO, SE5.E5_DATA, SE5.R_E_C_N_O_ REGSE5,
                E1_FILORIG, E1_EMISSAO, E1_PREFIXO, E1_NUM, E1_TIPO, E1_PARCELA, E1_CLIENTE, E1_LOJA, E1_XCHAVE, SE1.R_E_C_N_O_ REGSE1
         FROM   %table:SE5%  SE5
         INNER  JOIN %table:SE1% SE1
            ON  SE1.E1_FILORIG = SE5.E5_FILORIG 
            AND SE1.E1_PREFIXO = SE5.E5_PREFIXO 
            AND SE1.E1_NUM     = SE5.E5_NUMERO 
            AND SE1.E1_TIPO    = SE5.E5_TIPO 
            AND SE1.E1_CLIENTE = SE5.E5_CLIFOR 
            AND SE1.E1_PARCELA = SE5.E5_PARCELA
         WHERE  SE5.E5_FILORIG = %exp:cFilZ90% 
            AND SE5.E5_RECPAG  = 'R' 
            AND SE5.E5_MOTBX   = 'LIQ' 
            AND SE5.E5_DATA    > %exp:xDataV% 
            AND SE5.E5_MSEXP   = ' ' 
            AND SE5.E5_SITUACA = 'C'
				//AND SE5.E5_NUMERO  = '01042021A'	//	REMOVER
         ORDER  BY E5_NUMERO, E5_PREFIXO, E5_PARCELA, E5_TIPO, E5_CLIFOR, E5_LOJA 
		EndSql
	EndIf
	aQuery := GetLastQuery()

Return


/*/{Protheus.doc} HBHBAre082()
Função para envio do cancelamento da baixa do Titulo
@author Rene julian
@version P12.1.17
@since 24/04/2020
@return 
@obs
/*/
Static Function HBAre082(cAlias,cEnvJson,cMsgJSon, cID, cLoja )
	Local aCampJson := {}   //Campos do cabeçalho
	Local aInfoJson := {}   //Informações do cabeçalho
	Local oObjJSON  := JsonObject():New()
	Local nX        := 0
	Local lRet      := .T.
	Local lCancel   := .T.

	aCampJson := {"ID","Loja"}
	aInfoJson := {cID, cLoja }
	For nX := 1 to Len( aCampJson )
		oObjJSON[aCampJson[nX]] := aInfoJson[nX]
	Next nX
	cEnvJson := FWJsonSerialize(oObjJSON)
	FreeObj(oObjJSON)

	//chamo a função de envio via Rest para o Arena
	lRet := HBenvJSon(cEnvJson,@cMsgJSon,.T.)

Return lRet

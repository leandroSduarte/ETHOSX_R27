#include "rwmake.ch"
#include "protheus.ch"

/*/{Protheus.doc} HB06A02
//TODO Baixa do Contas a Pagar
@author Ethosx - MOA
@since 01/08/2020
@version 1.0
@return ${return}, ${return_description}

@type function
/*/
User Function HB06A02(aCompany)

	//Local   aCompany	:= {  }
	Local  	aRecs   	:= {  }
	Local   aExec   	:= Array(08)

	Local  	 cAlias  	:= ""
	Local   cWhere  	:= ""
	Local   cType   	:= ""
	Local   cPrefix 	:= ""
	Local   cBank   	:= ""
	Local   cAg     	:= ""
	Local   cAccount	:= ""
	Local   cNature 	:= ""
	Local   cAux    	:= ""
	Local   cChave  	:= ""
	Local 	cHistBaixa	:= "Baixas Pagar Autom."

	Local  	dBegin
	Local	dVenc

	Local	nX      	:= 0

	Private cFunName    := "HB06A02"
	Private lMSErroAuto := .F.

	Default aCompany 	:= {"01","0001010001"}

	RPCSetType(3)
	If RPCSetEnv(aCompany[01],aCompany[02],,,"FIN")
		ConOut("["+cFunName+"] Iniciando processamento... ["+FWTimeStamp(2)+"]")

		cType   := AllTrim(SuperGetMV("HB_06A02TP",,"UNI"         ))
		cPrefix := AllTrim(SuperGetMV("HB_06A02PR",,"*"         ))
		cNature := AllTrim(SuperGetMV("HB_06A02NT",,"OUTROS"  ))
		// Ethosx - MOA - 16/04/2021
		// Alteração para padrão novo - Conta Cofre.
	/*
	cBank   := AllTrim(SuperGetMV("HB_06A02BC",,"CX1"      ))
	cAg     := AllTrim(SuperGetMV("HB_06A02AG",,"00001"     ))
	cAccount:= AllTrim(SuperGetMV("HB_06A02AC",,"0000000001"))
	*/
		cCxLoja := SuperGetMV("HB_06A02D",,"COF/000001/"+Space(TamSX3("A6_NUMCON")[01]))

		cBank  	:= SUBS(cCxLoja,1,TamSX3("A6_COD")[01])
		cAg		:= SUBS(cCxLoja,AT("/",cCxLoja)+1,TamSX3("A6_AGENCIA")[01])
		cAccount:= AllTrim(SUBS(cCxLoja,RAT("/",cCxLoja)+1,TamSX3("A6_NUMCON")[01]))

		cHist	:= "Baixa automatica HB06A02 "+DtoC(Date())

		//dBegin  := STOD(SuperGetMV("HB_06A02BG",,"20210101"))
		dBegin  := dDataBase-30

		cWhere := "%"
	/*
		If cType <> "*"
		cWhere += " AND E2_TIPO IN "+FormatIn(cType,";")
		Else
		cWhere += " AND E2_TIPO <> ('PA')"
		EndIf
	*/
/*
	cWhere += " AND E2_TIPO = 'UNI' 

		If cPrefix <> "*"
		cWhere += " AND E2_PREFIXO IN "+FormatIn(cPrefix,";")
		EndIf
*/
		//cWhere += " AND E2_PREFIXO = 'UNI' "


		// cWhere += "AND E2_VENCREA >= '"+DTOS(dBegin)+"'%"  // DESCOMENTAR PARA EXECUÇÃO MANUAL

		dVenc := dBegin  // DESCOMENTAR PARA EXECUÇÃO MANUAL
		dDtProc := dDatabase

		For nX := 1 to DateDiffDay( dVenc , dDtProc ) // DESCOMENTAR PARA EXECUÇÃO MANUAL

			cAlias := GetNextAlias(  )
			If Select(cAlias) > 0
				(cAlias)->(DBCloseArea(  ))
			EndIf

			BeginSQL Alias cAlias
			SELECT
			SE2.R_E_C_N_O_ as RECSE2,
			E2_FILIAL, E2_PREFIXO, E2_NUM,  E2_PARCELA,
			E2_FORNECE, E2_LOJA, E2_TIPO, E2_EMISSAO,
			E2_VENCTO, E2_VENCREA, E2_FILORIG,
			E2_VALOR, E2_SALDO, E2_NATUREZ
			FROM
			%Table:SE2% (NOLOCK) SE2
			WHERE
			//E2_FILIAL = %xFilial:SE2%
			E2_FILIAL <>''
			AND E2_PREFIXO = 'UNI'
			AND E2_SALDO > 0
			AND E2_VENCREA = %Exp:DTOS(dVenc)%	
			AND SE2.%NotDel%
			EndSQL

			DBSelectArea(cAlias)
			If (cAlias)->(EOF(  ))
				ConOut("["+cFunName+"] Nenhum titulo encontrado para realizar a baixa em lote...")
				ConOut(GetLastQuery()[2])
				//Alert("["+cFunName+"] Nenhum titulo encontrado para realizar a baixa em lote... ["+FWTimeStamp(2)+"]")
			Else
				Do while (cAlias)->(!EOF(  ))
					aAdd(aRecs, (cAlias)->RECSE2)
					(cAlias)->(DBSkip(  ))

				End Do
			EndIf

			If Len(aRecs) > 0

				dDatabase :=  dVenc

				aExec[1] := aRecs
				aExec[2] := PAD(cBank,TAMSX3('A6_COD')[1])
				aExec[3] := PAD(cAg,TAMSX3('A6_AGENCIA')[1])
				aExec[4] := PAD(cAccount,TAMSX3('A6_NUMCON')[1])
				aExec[5] := PAD("",TAMSX3('EF_NUM')[1])
				aExec[6] := PAD("",TAMSX3('E2_LOTE')[1])
				aExec[7] := PAD(cNature,TAMSX3('ED_CODIGO')[1])
				aExec[8] := dVenc

				Begin Transaction
					MSExecAuto({|x,y| Fina090(x,y)},3,aExec)

					If lMSErroAuto
						DisarmTransaction(  )
						ConOut("["+cFunName+"] Ocorreu um problema ao realizar a baixa dos titulos em lote... ["+FWTimeStamp(2)+"]")
						ConOut(Mostraerro())
						// Alert("["+cFunName+"] Ocorreu um problema ao realizar a baixa dos titulos em lote... ["+FWTimeStamp(2)+"]")

					Else
						ConOut("["+cFunName+"] "+CVALTOCHAR( Len(aRecs) )+" titulo(s) baixados com sucesso! Lote: "+SE5->E5_LOTE+"... ["+FWTimeStamp(2)+"]")
					EndIf

				End Transaction
			EndIf

			dVenc := dBegin + nX  // DESCOMENTAR PARA EXECUÇÃO MANUAL

		Next nX  // DESCOMENTAR PARA EXECUÇÃO MANUAL

		ConOut("["+cFunName+"] Fim processamento... ["+FWTimeStamp(2)+"]")
		// alert("["+cFunName+"] Fim processamento... ["+FWTimeStamp(2)+"]")

		(cAlias)->(DBCloseArea(  ))

		RPCClearEnv(  )
	EndIf

Return Nil


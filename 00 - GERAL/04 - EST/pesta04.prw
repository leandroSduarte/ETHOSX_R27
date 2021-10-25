#INCLUDE "RWMAKE.CH"

/*/{Protheus.doc} PESTA04
//TODO Rotina que estorna o ultimo fechamento de estoque.
@author Carlos R. Moreira 
@since 26/05/2004
@version 1.2
@return Nil

@obs 		Adicionado o filtro de filial "aSelFil"  -  Ethosx - MOA - 19/06/2020.

@type function
/*/
User Function PESTA04()

	Local aSays     		:= {}
	Local aButtons  	:= {}
	Local aSelFil 			:= {}

	Local cCadastro	:= OemToAnsi("Estorna o Ultimo Fechamento de Estoque")

	Local nOpca     		:= 0

	// Private  cArqTxt

	aSelFil := 	AdmGetFil(.F.,.F.,"SB9")
	If Len( aSelFil ) <= 0
		Return
	EndIf



	// Pergunte(cPerg,.F.)

	Aadd(aSays, OemToAnsi(" Este programa ira processar os arquivos de fechamento de Estoque das Filiais "))
	Aadd(aSays, OemToAnsi(" selecionadas, para estornar o ultimo." ))

	Aadd(aButtons, { 1, .T., { || nOpca := 1, FechaBatch()  }})
	Aadd(aButtons, { 2, .T., { || FechaBatch() }})

	FormBatch(cCadastro, aSays, aButtons)

	If nOpca == 1

		Processa( { || ProcEstorno(aSelFil) }, "Processando o Estorno . . .")  //

	EndIf

Return

/*/{Protheus.doc} ProcEstorno
//TODO Processa estorno do o ultimo fechamento de estoque.
@author Carlos R. Moreira 
@since 26/05/2004
@version 1.2
@return Nil

@type function
/*/
Static Function ProcEstorno(aSelFil)

	Local aAreaSM0 	:= SM0->(GETAREA())
	Local cFilBkp 		:= cFilAnt
	Local nCont 			:= 0
	Local dULMes

	If MsgYesNo("Confirma o estorno do ultimo Fechamento das filiais selecionadas? " )
		ProcRegua(Len(aSelFil))

		For nCont := 1 to Len(aSelFil)
			
			IncProc("Estornando o Fechamento ... "+cValToChar(  Round((nCont / Len(aSelFil))*100,2)  )+"%" )

			SM0->(DbGoTop ())
			SM0->(MsSeek (cEmpAnt+aSelFil[nCont],.T.))

			cFilAnt := SM0->M0_CODFIL
			dULMes  := SuperGetMV("MV_ULMES",,,cFilAnt) 

			DbSelectArea("SB9")
			DbSetOrder(1)
			DbSeek(xFilial("SB9"))

			While SB9->(!Eof()) .And. xFilial("SB9") == SB9->B9_FILIAL 
				If SB9->B9_DATA # dUlMes
					DbSkip()
					Loop                                          
				EndIf    

				DbSelectArea("SB9")
				RecLock("SB9",.F.)
				SB9->(DbDelete())
				MsUnlock()
				DbSkip()

			End
			If DateDiffMonth(Date(  ), LastDate(MonthSub(dULMes,1)) ) <= 2
				PutMV('MV_ULMES', LastDate(MonthSub(dULMes,1)) )
			EndIf
				
		NEXT
	EndIf

	cFilAnt := cFilBkp
	RESTAREA(aAreaSM0)

Return 

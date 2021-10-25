#include 'protheus.ch'
#include 'parmtype.ch'

User Function ConvCBar()     

	Local cCodBar	:=""

	If !EMPTY(SE2->E2_CODBAR)
		cCodBar := IF(LEN(AllTrim(SE2->E2_CODBAR))<44,AllTrim(SE2->E2_CODBAR)+REPL("0",47-LEN(AllTrim(SE2->E2_CODBAR))),AllTrim(SE2->E2_CODBAR))
	EndIf

	If !IsBlind( ) .AND. Empty(cCodBar)
		MsgAlert("A linha digitável deve possuir no máximo 47 digitos (sem espaços e pontos) para a correta conversão do codigo de barras."+Chr(10)+chr(13),"Atenção")
		Return()
	EndIf
    /*
   	If !IsBlind( ) .AND. Len(cCodBar) > 47
    	MsgAlert("A linha digitável deve possuir no máximo 47 digitos (sem espaços e pontos) para a correta conversão do codigo de barras."+Chr(10)+chr(13),"Atenção")
		Return()
	EndIf */

	Do Case
		Case Len(cCodBar) == 47
		cCodBar := SUBSTR(cCodBar,1,4)+SUBSTR(cCodBar,33,15)+SUBSTR(cCodBar,5,5)+SUBSTR(cCodBar,11,10)+SUBSTR(cCodBar,22,10)

		Case Len(cCodBar) == 48
	   //	cCodBar := SUBSTR(cCodBar,1,48)
	    cCodBar := SUBSTR(cCodBar,1,11)+SUBSTR(cCodBar,13,11)+SUBSTR(cCodBar,25,11)+SUBSTR(cCodBar,37,11) 
	    
		Otherwise
		cCodBar := cCodBar+SPACE(48-LEN(cCodBar))

	EndCase

Return cCodBar


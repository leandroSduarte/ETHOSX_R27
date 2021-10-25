#INCLUDE "RWMAKE.CH"
#INCLUDE "PROTHEUS.CH"
/*
ÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜÜ
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
±±ÉÍÍÍÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍËÍÍÍÍÍÍÑÍÍÍÍÍÍÍÍÍÍÍÍÍ»±±
±±ºPrograma  ³HABGCODBARºAutor  ³Eduardo Ramalho     º Data ³  08/02/2019 º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÊÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºDesc.     ³Transforma a linha digitavel em codigo de barras.           º±±
±±º          ³                                                            º±±
±±ÌÍÍÍÍÍÍÍÍÍÍØÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¹±±
±±ºUso       ³ Gatilho no campo E2_CODBAR                                 º±±
±±ÈÍÍÍÍÍÍÍÍÍÍÏÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍÍ¼±±
±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±±
ßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßßß
*/
User Function HABGCODBAR(_cCodbar)

Local _cMonta1 := ""
Local _cMonta2 := ""
Local _cMonta3 := ""
Local _cMonta4 := ""
Local _cMonta5 := ""
Local _cMonta6 := ""
Local _cMonta7 := ""
Local _cReturn := ""
Local _nTamCod := 0
Local _nTamVal := 0
                               
Default _cCodbar := AllTrim(M->E2_CODBAR)

_nTamCod := Len(_cCodbar)
_nTamVal := Len(_cCodbar)-33


// Verifica tamanho da linha digitável

If _nTamCod == 44
	_cReturn := _cCodbar
Else
	If _nTamVal < 14        // Codigo de Barras sem fator de vencimento.
		_cMonta3 := "0000"
		_cMonta4 := Strzero(Val(Substr(_cCodbar,34,_nTamVal)),10)
	Else                    // Codigo de Barras com fator de vencimento.
		_cMonta3:= Substr(_cCodbar,34,4)
		_cMonta4:= Substr(_cCodbar,38,10)
	Endif
	
	// Monta o Codigo de Barras
	
	_cMonta1:= Substr(_cCodbar,1,4)
	_cMonta2:= Substr(_cCodbar,33,1)
	_cMonta5:= Substr(_cCodbar,5,5)
	_cMonta6:= Substr(_cCodbar,11,10)
	_cMonta7:= Substr(_cCodbar,22,10)
	
	_cCodBar := _cMonta1 + _cMonta2 + _cMonta3 + _cMonta4 + _cMonta5 + _cMonta6 + _cMonta7
	_cReturn := _cCodbar
Endif
Return(_cReturn)

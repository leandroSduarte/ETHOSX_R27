#Include "Protheus.ch"
/*
//Rotina que trata o trava do m s senj· financeiro Fiscal ou Contabil
*/
User Function GgPrmCtb()
Local oDlg

Local _cUsrCtb	:= GetMv("MV_XUSRCTB")
Local _cUsrFin	:= GetMv("MV_XUSRFIN")
Local _cUsrFis	:= GetMv("MV_XUSRFIS")
Local _cDtCtb	:= ""
Local _cDtFin	:= ""
Local _cDtFis	:= ""
Local _nLin		:= 0

Local _cUsr		:= RetCodUsr()

nOpca := 0

If ( _cUsr $ _cUsrCtb )  .Or. ( _cUsr $ _cUsrFin )  .Or. ( _cUsr $ _cUsrFis )
	
	DEFINE MSDIALOG oDlg TITLE OemToAnsi("Par‚metros") FROM  15,6 TO 190,366 PIXEL OF oMainWnd

    _nLin	:= -10
	if _cUsr $ _cUsrCtb 
		_cDtCtb	:= GETMV('MV_ULMES',,CTOD('01/01/2006'))
		_nLin +=20
		@_nLin,10 Say "MV_ULMES:" Pixel Of oDlg
		@_nLin,50 Get _cDtCtb Size 50,10 Picture "@!" Pixel Of oDlg
    Endif
	if _cUsr $ _cUsrFin
		_cDtFin	:= GETMV('MV_DATAFIN',,CTOD('31/12/2006'))
		_nLin +=20
		@_nLin,10 Say "MV_DATAFIN:"  Pixel Of oDlg
		@_nLin,50 Get _cDtFin Size 50,10 Picture "@!" Pixel Of oDlg
	Endif
	if _cUsr $ _cUsrFis
		_cDtFis	:= GETMV('MV_DATAFIS',,CTOD('31/12/2006'))
		_nLin +=20
		@_nLin,10 Say "MV_DATAFIS:"  Pixel Of oDlg
		@_nLin,50 Get _cDtFis Size 50,10 Picture "@!" Pixel Of oDlg
	Endif

	
	DEFINE SBUTTON FROM 71,124 TYPE 1 ENABLE OF oDlg ACTION (nOpca := 1,oDlg:End())
	DEFINE SBUTTON FROM 71,152 TYPE 2 ENABLE OF oDlg ACTION (nOpca := 0,oDlg:End())
	
	ACTIVATE MSDIALOG oDlg CENTERED
	
	If nOpca == 1
	
		if _cUsr $ _cUsrCtb 
			PutMv("MV_ULMES",_cDtCtb )   
		Endif
		if _cUsr $ _cUsrFin
			PutMv("MV_DATAFIN",_cDtFin )
		Endif
		if _cUsr $ _cUsrFis
			PutMv("MV_DATAFIS",_cDtFis )
		Endif
	EndIf
Else
	Aviso("Aviso","Usu·rio se?m acesso a rotina" ,{"Ok"})
	Return
EndIf
Return()

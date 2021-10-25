#Include 'TOTVS.ch'

//-----------------------------------------------------------------------
/*{Protheus.doc} REX06A01
@Description Importa um arquivo CSV e gera os titulos a pagar referente
as provisoes.
@return             Nenhum
@author             Leonardo Espinosa
@since              14/10/2020
@version            1.0
@project            HABIBS
*/
//-----------------------------------------------------------------------
User Function REX06A01( )
    Local       aParamBox   := {}
    Local       aRet        := {}
	Local       cFile	    :=	""
    Local       nOpc        := 0

    Private     oProcess

	/*Variáveis para tela de log*/
	Private	oSize		:=	FwDefSize():New()
	Private aObjects	:={100,100,.T.,.T.}
	Private aInfo		:=	{oSize:aWindSize[1],oSize:aWindSize[2],oSize:aWindSize[3],oSize:aWindSize[4],5,5}
	Private	oFont		:= TFont():New('Courier new',,-12,.T.)
	Private	aButtons	:=	{} 

    aAdd(aParamBox,{3,"Operacao",1,{"Incluir","Excluir"},50,"",.T.})
    If ParamBox(aParamBox,"Importador - Contas a pagar...",@aRet)

        cFile:= cGetFile("Arquivos CSV  (*.CSV)  | *.CSV  ","",1,"C:\",.T.,GETF_LOCALHARD,.T.,.T.)

        If !Empty(cFile)
            If aRet[1] == 1 //Inclusao
                nOpc    := 3
            Else 
                nOpc    := 5
            EndIf

            oProcess := MsNewProcess():New({|lEnd| ProcArq(cFile, @oProcess, @lEnd, nOpc) },"Processamento","Criacao de Provisoes",.T.)
            oProcess:Activate()

        EndIf
    EndIf 

Return

//-------------------------------------------------------------
// Processa o arquivo de acordo com os parametros selecionados
//-------------------------------------------------------------
Static Function ProcArq(cFile, oProcess, lEnd, nOpc)
    Local AreaPrc 			:= GetArea()
	Local aLinProc			:={}
	Local aCab 			    :={}
    Local aErroAuto         :={}

	Local cNomeUsr		    := Upper(UsrRetName(RetCodUsr()))
	Local cCodFor           := ""
    Local cLjFor            := ""
    Local cMsgLog           := ""
    Local cLogErro          := ""
    Local cParcela          := ""

    Local dDtBaseBkp		:= dDataBase
    
    Local lErro             := .F. 

    Local nCount            := 0
	Local nLinCount		    := 1
	Local nAux 				:= 0
    Local nLinProc			:= 0
	Local nLinOk			:= 0
	Local nLinErro			:= 0
	Local nContaErro		:= 0
	Local nCalenErro		:= 0
	
    PRIVATE lMsErroAuto := .F. 
    Private lMsHelpAuto	:= .T. 
    Private lAutoErrNoFile := .T. 

    If !File(cFile)
        MsgStop("Atenção"+CRLF+;
            "Arquivo invalido. Verifique o parametro de arquivo!", "Importacao de Provisoes")

        Return
    
    EndIf

    FT_FUSE(cFile)
	nRegArq := FT_FLastRec()
	FT_FGOTOP()

	oProcess:SetRegua1(nRegArq)

	cNumLin	:= "001"
    cMsgLog := ""

	lCab    :=	.T.

	While !FT_FEOF()
		nAux++

		If nAux > nRegArq
			Exit
		EndIf
		oProcess:IncRegua1("Linha processada:" + cValToChar(nLinCount))
		cLinha		:= FT_FREADLN()
		aLinProc	:=	Separa(cLinha,";",.T.)

		If  Valtype(aLinProc) == 'A'  .AND. Len(aLinProc) > 0
			lErro := .F. 
            cFilialPrc := Padl(AllTrim(aLinProc[1]),TamSX3("CT2_FILIAL")[1],"0")

			If	Upper(aLinProc[1])	!="FILIAL" .And. FwFilExist(cEmpAnt,cFilialPrc)
                nLinProc++

                cFilAnt := cFilialPrc

                cCodFor := PADL(aLinProc[07],TamSX3("A2_COD")[1]    ,"0")
                cLjFor  :=aLinProc[08]

                cParcela := IIF(IsDigit(aLinProc[04]),STRZERO(VAL(aLinProc[04]),TamSX3("E2_PARCELA")[01]),aLinProc[04])

                DBSelectArea("SED")
                DBSetOrder(1)
                If SED->(!DBSeek(xFilial("SED")+aLinProc[06]))
                    lErro := .T. 

                    cLogErro := " NATUREZA "+aLinProc[06]+" INVALIDA "

                    cMsgLog += "Linha "+cValToChar(nLinProc)+" | Filial:"+aLinProc[01]+"| Prefixo: "+aLinProc[02]+"| Titulo "+aLinProc[03]+" - Erro ao incluir o titulo:"
                    cMsgLog += CRLF+cLogErro
                EndIf 

                DBSelectArea("SA2")
                DBSetOrder(1)
                If SA2->(!DBSeek(xFilial("SA2")+cCodFor+cLjFor)) .AND. IIF(FieldPos("A2_MSBLQL")>0,SA2->A2_MSBLQL <> "1",.T.)
                    lErro := .T. 

                    cLogErro := " Fornecedor "+cCodFor+"/"+cLjFor+" INVALIDO OU BLOQUEADO "

                    cMsgLog += "Linha "+cValToChar(nLinProc)+" | Filial:"+aLinProc[01]+"| Prefixo: "+aLinProc[02]+"| Titulo "+aLinProc[03]+" - Erro ao incluir o titulo:"
                    cMsgLog += CRLF+cLogErro
                Endif
                If !lErro
                    DBSelectArea("SE2")
                    DBSetOrder(1)
                    // If SE2->(!DBSeek(xFilial("SE2")+aLinProc[02]+aLinProc[03]+cParcela+aLinProc[05]+cCodFor+cLjFor))
                        dDataBase   := CTOD(aLinProc[10])

                        aCab := {   { "E2_FILIAL"   , STRZERO(VAL(aLinProc[01]),TamSX3("E2_FILIAL")[1])                      , NIL },;
                                    { "E2_FILORIG"  , STRZERO(VAL(aLinProc[01]),TamSX3("E2_FILIAL")[1])                      , NIL },;
                                    { "E2_PREFIXO"  , aLinProc[02]                      , NIL },;
                                    { "E2_NUM"      , aLinProc[03]                      , NIL },;
                                    { "E2_PARCELA"  , cParcela                          , NIL },;
                                    { "E2_TIPO"     , aLinProc[05]                      , NIL },;
                                    { "E2_NATUREZ"  , aLinProc[06]                      , NIL },;
                                    { "E2_FORNECE"  , cCodFor                           , NIL },;
                                    { "E2_LOJA"     , cLjFor                            , NIL },;
                                    { "E2_EMISSAO"  , CTOD(aLinProc[10])                , NIL },;
                                    { "E2_VENCTO"   , CTOD(aLinProc[11])                , NIL },;
                                    { "E2_VENCREA"  , DataValida(CTOD(aLinProc[11]))    , NIL },;
                                    { "E2_VALOR"    , Val(STRTRAN(STRTRAN(STRTRAN(aLinProc[12],".",""),",","."),"R$","")), NIL },;
                                    { "E2_HIST"     , NoAcento(AllTrim(aLinProc[13]))   , NIL } }

                        cLogErro := ""
                        aErroAuto := {}
                        aSize(aErroAuto, Len(aErroAuto))

                        lMsErroAuto := .F. 
                        MsExecAuto( { |x,y,z| FINA050(x,y,z)}, aCab,, nOpc) 
                        
                        If lMsErroAuto
                            aErroAuto := GetAutoGRLog()
                            For nCount := 1 To Len(aErroAuto)
                                cLogErro += AllTrim(aErroAuto[nCount])+CRLF

                            Next nCount

                            cMsgLog += "Linha "+cValToChar(nLinProc)+" | Filial:"+aLinProc[01]+"| Prefixo: "+aLinProc[02]+"| Titulo "+aLinProc[03]+" - Erro ao incluir o titulo:"
                            cMsgLog += CRLF+cLogErro

                        
                        Else
                            cMsgLog += "Linha "+cValToChar(nLinProc)+" | Filial:"+aLinProc[01]+"| Prefixo: "+aLinProc[02]+"| Titulo "+aLinProc[03]+" - Titulo Incluido com sucesso!."
                        Endif
                    // Else
                    //     cMsgLog += "Linha "+cValToChar(nLinProc)+" | Filial:"+aLinProc[01]+"| Prefixo: "+aLinProc[02]+"| Titulo "+aLinProc[03]+" - Titulo ja existente no contas a pagar."
 
                    // EndIf 
                // Else
                //     cMsgLog += "Linha "+cValToChar(nLinProc)+" | Filial:"+aLinProc[01]+"| Prefixo: "+aLinProc[02]+"| Titulo "+aLinProc[03]+" - Fornecedor nao encontrado ou bloqueado."

                EndIf

            EndIf

            cMsgLog += CRLF+CRLF+Replicate("=",50)+CRLF
        
        EndIf

		aCab	:={}
		aSize(aCab,len(aCab))

		nLinCount++

		FT_FSKIP()

		oProcess:SetRegua2(nLinCount)

	EndDo

    If Len(cMsgLog) > 0
        EECVIEW(cMsgLog)
    EndIf 

    dDataBase	:= dDtBaseBkp

Return

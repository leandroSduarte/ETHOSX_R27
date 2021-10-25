#Include "TOTVS.ch"
#Include "FWMVCDEF.ch"
#INCLUDE "PROTHEUS.CH"    
#INCLUDE "TOPCONN.CH" 
#include "rwmake.ch"  
#include "fileio.ch"    
#INCLUDE "FWPrintSetup.ch"
#INCLUDE "RPTDEF.CH"
#Include "DBTREE.CH"
#Include "HBUTTON.CH"
#Define XENTERX Chr(13)+Chr(10) 
//========================================================================================================================================   
//Programa............: ALLCADEMP()
//Autor...............: Paulo César (PC) 
//Data................: 25/05/2019
//Descricao / Objetivo: Relação de Prefixo X Empresa
//Cliente             : ETHOS X - HABBIBS
//============================================================================================================
User Function ALLCADEMP()
//=============================================================
Private axArea      := GetArea()
Private oBrowse     := FwLoadBrw("ALLCADEMP")
Private aRotina     := FwMVCMenu("ALLCADEMP")//MenuDef()
Private cCadastro   := "Relação de Prefixo (CSV)X Empresa"
Private lMsErroAuto := .T.
Private lMarcar  	:= .F.
Private axLinha     := {}
Private cQryAux     := ""

oMark := FWMarkBrowse():New()
oMark:SetAlias('SX5')
SET FILTER TO ALLTRIM(SX5->X5_TABELA) == "Z3"
oMark:Activate()
RestArea( axArea )
Return
//=============================================================
Static Function BrowseDef()
//=============================================================
Local oBrowse := FwMBrowse():New()

oBrowse:SetAlias("SX5")
oBrowse:SetDescription("Relação de Prefixo X Empresa")
oBrowse:SetMenuDef("ALLCADEMP")
Return (oBrowse)

//=============================================================
Static Function ModelDef()
//=============================================================
Local oModel := MPFormModel():New("PAOLLAM",,)
Local oStruSX5 := FwFormStruct(1, "SX5")

oModel:AddFields("SX5MASTER", NIL, oStruSX5)
oModel:SetPrimaryKey({'SX5_FILIAL'})
oModel:SetDescription("Relação de Prefixo X Empresa")
oModel:GetModel("SX5MASTER"):SetDescription("Relação de Prefixo X Empresa")
Return (oModel)

//=============================================================
Static Function ViewDef()
//=============================================================
Local nXtamX := 100
Local oView := FwFormView():New()
Local oStruSX5 := FwFormStruct(2, "SX5")
Local oModel := FwLoadModel("ALLCADEMP")

oView:SetModel(oModel)
oView:AddField("VIEW_SX5", oStruSX5, "SX5MASTER")
oView:CreateHorizontalBox("SUPERIOR", nXtamX)
oView:SetOwnerView("VIEW_SX5", "SUPERIOR")
Return (oView)

//=======================================================================================================================================
Static Function MenuDef()
//=======================================================================================================================================
   Local aRotina := {}  //FwMVCMenu("ALLCADEMP")
   
ADD OPTION aRotina TITLE "Visualizar"       ACTION "VIEWDEF.ALLCADEMP"        OPERATION  2 ACCESS 0
ADD OPTION aRotina TITLE "Incluir"          ACTION "VIEWDEF.ALLCADEMP"        OPERATION  3 ACCESS 0
ADD OPTION aRotina TITLE "Alterar"          ACTION "VIEWDEF.ALLCADEMP"        OPERATION  4 ACCESS 0
ADD OPTION aRotina TITLE "Excluir"          ACTION "VIEWDEF.ALLCADEMP"        OPERATION  5 ACCESS 0
Return (aRotina)

//====================================================================================
Static Function ZRetLog(aErr,cLit)
//====================================================================================

Local lHelp   := .F.
Local lTabela := .F.
Local cLinha  := ""
Local aRet    := {}
Local nI      := 0

For nI := 1 to LEN(aErr)
	cLinha  := UPPER( aErr[nI] )
	cLinha  := STRTRAN( cLinha,CHR(13), " " )
	cLinha  := STRTRAN( cLinha,CHR(10), " " )	
	If SUBS( cLinha, 1, 4 ) == 'HELP'
		lHelp := .T.
	EndIf	
	If SUBS( cLinha, 1, 6 ) == 'TABELA'
		lHelp   := .F.
		lTabela := .T.
	EndIf	
	If  lHelp .or. ( lTabela .AND. '< -- INVALIDO' $  cLinha )
		aAdd( aRet,  cLinha )
	EndIf	
Next
Return aRet


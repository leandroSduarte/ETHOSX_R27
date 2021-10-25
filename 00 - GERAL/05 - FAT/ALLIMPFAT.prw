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
//-------------------------------------------------------------------------  
//Programa............: ALLIMPFAT()  
//Autor...............: Paulo César (PC)  
//Data................: 25/05/2019
//Descricao / Objetivo: Ajusta Dicionário de Dados 
//Cliente             : ETHOS X - HABBIBS
//-------------------------------------------------------------------------
User Function ALLIMPFAT()
//-------------------------------------------------------------------------
Private aTitulo     := {}
Private cCadastro   := "Importação de Pedidos de Venda"
Private oBrowse     := FwLoadBrw("ALLIMPFAT")
Private aRotina     := {}
Private lMsErroAuto := .T.
Private lMarcar  	:= .F.
Private axLinha     := {}
Private cEmpX       := ""   
Private cFilX       := ""  
Private cPrefxArq   := ""
Private bKeyF12	    //:= {||  oMark:SetInvert(.F.),oMark:Refresh(),oMark:GoTop(.T.) } //Programar a tecla F12
Private cEmpBKP     := cEmpAnt
Private cFilBKP     := cFilAnt
Private dDBaseBKP   := dDataBase
Private aArqPPM     := {}
Private cPathTmp    := ""
Private aSM0Stru    :=  {}
Private cArqLog     := "LOGCSV.LOG"
Private cFileC      := ""
Private cLog        := ""
Private cStartPath  := GetSrvProfString("Startpath","")
Private nTotAdm     := 0
Private nTotSrv     := 0
Private nTotMid     := 0
Private oMark
Private oMsg
Private cMsg := ""
Private oDlgNFS
Private oDlgNFC
Private oWBrowse1
Private aWBrowse1 := {}
Private lFirst := .T.
Private cMVFATSER :=  alltrim(GetNewPar('MV_FATSER','')) 
Private cMVRoyalt :=  alltrim(GetNewPar('HB_ROYALTI','')) 
Private aMVFATSER := {}
Private aMVRoyalt := {}
Private nTotal    := 0
Private cPrefxTel   := ""
Private cQryAux     := ""

cMVFATSER := '{"'+Alltrim(cMVFATSER)+'"}'
cMVFATSER := StrTran(cMVFATSER,';','","')//adiciona o cLinha no array trocando o delimitador ; por , para ser reconhecido como elementos de um array 
aMVFATSER := aClone(&(cMVFATSER)) 

cMVRoyalt :=  alltrim(GetNewPar('HB_ROYALTI','')) 
//cMVRoyalt := '{'+Alltrim(cMVRoyalt)+'}'
cMVRoyalt := StrTran(cMVRoyalt,";",",")//adiciona o cLinha no array trocando o delimitador ; por , para ser reconhecido como elementos de um array 
aMVRoyalt := aClone(&(cMVRoyalt)) 

cLog := "UPDATE "+RetSqlName("PPX")+" SET PPX_OK = '  '"
TCSQLExec(cLog)

cLog := "DELETE FROM "+RetSqlName("PPX")+" WHERE D_E_L_E_T_ = '*'"
TCSQLExec(cLog)
cLog := "DELETE FROM "+RetSqlName("PPZ")+" WHERE D_E_L_E_T_ = '*'"
TCSQLExec(cLog)
cLog := "DELETE FROM "+RetSqlName("SED")+" WHERE D_E_L_E_T_ = '*'"
TCSQLExec(cLog)
cLog := "DELETE FROM "+RetSqlName("SE4")+" WHERE D_E_L_E_T_ = '*'"
TCSQLExec(cLog)

cPrefxTel := U_ETX_PREFIXO()
aRotina   := MenuDef()
aSM0Stru := SM0->(DBSTRUCT())
cCadastro   := "Importação de Pedidos de Venda"

fPrincipal()
cEmpAnt := cEmpBKP  
cFilAnt := cFilBKP 
OpenFile(cEmpAnt+cFilAnt)
OpenSM0()
Return

///------------------------------------------------------------------------- 
Static Function fPrincipal()
///-------------------------------------------------------------------------
fLegFat("")
aTitulo := U_ETX_EMPRE("PP")
oMark := FWMarkBrowse():New()
bKeyF12	:= {||  oMark:SetInvert(.F.),oMark:Refresh(),oMark:GoTop(.T.) } //Programar a tecla F12
oMark:SetAlias('PPX')
//SET FILTER TO PPX->PPX_PREFIX = Alltrim(cPrefxTel)
oMark:SetDescription("Importação de Pedidos de Venda")
oMark:SetFieldMark( 'PPX_OK' )
oMark:SetAllMark( { || oMark:AllMark() } )
oMark:AddLegend("PPX_STATUS=='0'","BR_VERMELHO"  ,"Importado com Erro" )
oMark:AddLegend("PPX_STATUS=='1'","BR_VERDE"   ,"Importado com sucesso" )
oMark:AddLegend("PPX_STATUS=='2'","BR_AZUL"      ,"Pedido de Venda gerado")
oMark:AddLegend("PPX_STATUS=='4'","BR_AZUL_CLARO","NF gerada parcial")
oMark:AddLegend("PPX_STATUS=='5'" ,"BR_LARANJA"  ,"NF gerada total")
oMark:bAllMark := { || U_ALL_INVERT (oMark:Mark(),lMarcar := !lMarcar ), oMark:Refresh(.T.)  }
oMark:Activate()
Return

//------------------------------------------------------------------------
Static Function fLegFat(cxID)
//------------------------------------------------------------------------
If Select("QRY_AUX")>0
   DbSelectArea("QRY_AUX")
   DbCloseArea()
Endif
cQryAux := " SELECT PPX_ID,PPX_STATUS,PPX_AGLUT"
cQryAux += ",ISNULL((SELECT COUNT(*) FROM "+RetSqlName("SC5")+" SC5 WHERE  SC5.D_E_L_E_T_ = ' ' AND LEFT(C5_OBS,6) = PPX_ID AND C5_NOTA <> ' '),0) PPX_COUNT"
cQryAux += " FROM "+RetSqlName("PPX")+" PPX"
cQryAux += " WHERE PPX.D_E_L_E_T_ = ' ' 
cQryAux += "AND PPX_STATUS IN ('4','5')"
If ALLTRIM(cxID)  <> ""
   cQryAux += "AND PPX_ID = '"+cxID+"'"
Endif
cQryAux += " ORDER BY PPX_ID"
cQryAux := If(cEmpAnt="99",Replace(cQryAux,"C*O*L*L*A*T*E*","COLLATE Latin1_General_CI_AS"),Replace(cQryAux,"C*O*L*L*A*T*E*",""))
cQryAux := ChangeQuery(cQryAux)
TCQuery cQryAux New Alias "QRY_AUX"
dbselectarea("QRY_AUX")                   
Count to nCount 
ProcRegua(nCount)
QRY_AUX->(DbGoTop())
While QRY_AUX->(!EOF())  
      dbSelectArea("PPX")
      PPX->(dbSetOrder(1))
      PPX->(dbSeek(xFilial("PPX")+QRY_AUX->PPX_ID))
      If PPX->PPX_STATUS $ "4|5"
         If PPX->PPX_AGLUT $ "G|S" 
            RecLock("PPX",.F.)
            If QRY_AUX->PPX_COUNT = 3
               PPX->PPX_STATUS = "5"
            Else
               PPX->PPX_STATUS = "4"
            Endif   
            PPX->(msUnlock())
         Endif  
         If !(PPX->PPX_AGLUT $ "G|S" )
            RecLock("PPX",.F.)
            If QRY_AUX->PPX_COUNT = 1
               PPX->PPX_STATUS = "5"
            Else
               PPX->PPX_STATUS = "4"
            Endif   
            PPX->(msUnlock())
         Endif   
      Endif          
      QRY_AUX->(dbSkip())
 Enddo
Return
//------------------------------------------------------------------------
User Function ALL_INVERT ()
//------------------------------------------------------------------------
Local cMarca    := oMark:Mark()
Local lInverte  := oMark:IsInvert()

dbSelectArea("PPX")
PPX->(dbGoTop())
While !PPX->(Eof())
       RecLock("PPX",.F.)
       PPX->PPX_OK := IIf(lMarcar,cMarca,'  ')
       PPX->(MsUnlock())
       PPX->(dbSkip())
EndDo
Return

//------------------------------------------------------------------------
Static Function BrowseDef()
//------------------------------------------------------------------------
Local oBrowse := FwMBrowse():New()

oBrowse:AddLegend("PPX_STATUS=='0'","BR_VERMELHO"  ,"Importado com Erro" )
oBrowse:AddLegend("PPX_STATUS=='1'","BR_VERDE"   ,"Importado com sucesso" )
oBrowse:AddLegend("PPX_STATUS=='2'","BR_AZUL"      ,"Pedido de Venda gerado")
oBrowse:AddLegend("PPX_STATUS=='4'","BR_AZUL_CLARO","NF gerada parcial")
oBrowse:AddLegend("PPX_STATUS=='5'" ,"BR_LARANJA"  ,"NF gerada total")
oBrowse:SetAlias("PPX")
oBrowse:SetDescription("Importação de Pedidos de Venda")
oBrowse:SetMenuDef("ALLIMPFAT")
Return (oBrowse)

///-------------------------------------------------------------------------
Static Function ModelDef()
///-------------------------------------------------------------------------
Local oModel := MpFormModel():New("PAOLLAM",{|oModel|fPreVld(oModel)},{|oModel|fVldSave(oModel)},{|oModel|fSave(oModel)})
Local oStruPPX := FwFormStruct(1, "PPX")
Local oStruPPZ := FwFormStruct(1, "PPZ")

oModel:AddFields("PPXMASTER", NIL, oStruPPX)
oModel:SetPrimaryKey({'PPX_FILIAL','PPX_ID'})
oModel:AddGrid("PPZDETAIL", "PPXMASTER", oStruPPZ)
oModel:SetRelation("PPZDETAIL", {{"PPZ_FILIAL", "FwXFilial('PPZ')"}, {"PPZ_ID", "PPX_ID"}}, PPZ->(IndexKey( 1 )))
oModel:SetDescription("Importação de Pedidos de Venda" )
oModel:GetModel("PPXMASTER"):SetDescription("Importação de Pedidos de Venda")
oModel:GetModel("PPZDETAIL"):SetDescription("Importação de Pedidos de Venda-Itens")
Return (oModel)

//-------------------------------------------------------------------------
Static Function ViewDef()
//-------------------------------------------------------------------------
Local oView := FwFormView():New()
Local oStruPPX := FwFormStruct(2, "PPX")
Local oStruPPZ := FwFormStruct(2, "PPZ")
Local oModel := FwLoadModel("ALLIMPFAT")

oView:SetModel(oModel)
oView:AddField("VIEW_PPX", oStruPPX,"PPXMASTER")
oView:AddGrid("VIEW_PPZ", oStruPPZ,"PPZDETAIL")
oView:CreateHorizontalBox("SUPERIOR", 60)
oView:CreateHorizontalBox("INFERIOR", 40)
oView:SetOwnerView("VIEW_PPX","SUPERIOR")
oView:SetOwnerView("VIEW_PPZ","INFERIOR")
oView:AddIncrementField("VIEW_PPZ","PPZ_LINE")
oView:EnableTitleView("VIEW_PPZ","Itens do Pedido")
Return (oView)

//------------------------------------------------------------------------
Static Function fPreVld(oModel)
//------------------------------------------------------------------------
Local lRet     := .T.
Local nOpc     := oModel:GetOperation()
Local cMsgErro := ""

lRet := FwFormCommit(oModel)
cMsgErro := ""
If PPX->PPX_STATUS $ "|2|3|4|5|"
   If oModel:nOperation > 2 .AND. oModel:nOperation  < 6
       If oModel:nOperation = 4
          If Alltrim(PPX->PPX_STATUS) = "2"
             cMsgErro += "Registro já gerado Pedido de Venda, não pode ser Alterado"+Chr(13)+Chr(10)
          Endif
          If Alltrim(PPX->PPX_STATUS) = "3"
             cMsgErro += "Registro já Excluído Pedido de Venda, não pode ser Alterado"+Chr(13)+Chr(10)
          Endif
          If Alltrim(PPX->PPX_STATUS) $ "|4|5|"
             cMsgErro += "Registro já gerado Nota Fiscal, não pode ser Alterado"+Chr(13)+Chr(10)
          Endif
       Endif  
   Endif
Endif
If Alltrim(cMsgErro)<> ""
   MsgStop(cMsgErro,"Erro1")
   lRet := .F. 
Endif   
Return (lRet)

//------------------------------------------------------------------------
Static Function fVldSave(oModel)
//------------------------------------------------------------------------
Local lRet     := .T.
Local nOpc     := oModel:GetOperation()
Local cMsgErro := ""

lRet := FwFormCommit(oModel)
cMsgErro := ""
If PPX->PPX_STATUS $ "|2|3|4|5|"
   If oModel:nOperation > 2 .AND. oModel:nOperation  < 6
       If oModel:nOperation = 4
          If Alltrim(PPX->PPX_STATUS) = "2"
             cMsgErro += "Registro já gerado Pedido de Venda, não pode ser Alterado"+Chr(13)+Chr(10)
          Endif
          If Alltrim(PPX->PPX_STATUS) = "3"
             cMsgErro += "Registro já Excluído Pedido de Venda, não pode ser Alterado"+Chr(13)+Chr(10)
          Endif
          If Alltrim(PPX->PPX_STATUS) $ "|4|5|"
             cMsgErro += "Registro já gerado Nota Fiscal, não pode ser Alterado"+Chr(13)+Chr(10)
          Endif
       Endif  
   Endif
Endif
If Alltrim(cMsgErro)<> ""
   MsgStop(cMsgErro,"Erro1")
   lRet := .F. 
Endif   
Return (lRet)

//------------------------------------------------------------------------
Static Function fSave(oModel)
//------------------------------------------------------------------------
Local lRet     := .T.
Local nOpc     := oModel:GetOperation()
Local cMsgErro := ""

lRet := FwFormCommit(oModel)
cMsgErro := ""
If PPX->PPX_STATUS $ "|2|3|4|5|"
   If oModel:nOperation > 2 .AND. oModel:nOperation  < 6
       If oModel:nOperation = 4
          If Alltrim(PPX->PPX_STATUS) = "2"
             cMsgErro += "Registro já gerado Pedido de Venda, não pode ser Alterado"+Chr(13)+Chr(10)
          Endif
          If Alltrim(PPX->PPX_STATUS) = "3"
             cMsgErro += "Registro já Excluído Pedido de Venda, não pode ser Alterado"+Chr(13)+Chr(10)
          Endif
          If Alltrim(PPX->PPX_STATUS) $ "|4|5|"
             cMsgErro += "Registro já gerado Nota Fiscal, não pode ser Alterado"+Chr(13)+Chr(10)
          Endif
       Endif  
   Endif
Endif
If Alltrim(cMsgErro)<> ""
   MsgStop(cMsgErro,"Erro1")
   lRet := .F. 
Endif   
Return (lRet)

//------------------------------------------------------------------------
Static Function MenuDef()
//------------------------------------------------------------------------
Local aRotina  := {}  //FwMVCMenu("ALLIMPFAT")     
Local aRotina2 := {}
Local aRotina3 := {}
Local aRotina4 := {} 
Local aRotina5 := {}
Local cPrefxTel := U_ETX_PREFIXO()

aAdd(aRotina,{"Visualizar","VIEWDEF.ALLIMPFAT",0,02})  
aAdd(aRotina,{"Incluir"   ,"VIEWDEF.ALLIMPFAT",0,03})
aAdd(aRotina,{"Alterar"   ,"VIEWDEF.ALLIMPFAT",0,04})
aAdd(aRotina,{"Excluir"   ,"VIEWDEF.ALLIMPFAT",0,05})

aAdd(aRotina2,{"Importar"         ,"U_ALL_EXEC('Importar')"      ,0,10})  
aAdd(aRotina2,{"Avaliar"          ,"U_ALL_EXEC('Avaliar')"       ,0,11})   
aAdd(aRotina2,{"Excl Imp em Lote" ,"U_ALL_EXEC('Lote')"          ,0,12})    
 

aAdd(aRotina3,{"Gerar PV"  ,"U_ALL_EXEC('Gerar')"                ,0,14})  
aAdd(aRotina3,{"Excluir PV"  ,"U_ALL_EXEC('Excluir PV')"         ,0,15}) 
aAdd(aRotina3,{"Excluir PV Lote","U_ALL_EXEC('Excluir PV Lote')" ,0,16})       

aAdd(aRotina4,{"Gerar NF"     ,"U_ALL_EXEC('NF')"                ,0,18})  
aAdd(aRotina4,{"Excluir NF"   ,"U_ALL_EXEC('EXCLUIR NF')"        ,0,19})  
aAdd(aRotina4,{"Transmitir NF","U_ALL_EXEC('Transmitir NF')"     ,0,20})
aAdd(aRotina4,{"Histórico"     ,"U_ALL_EXEC('Historico NF')"     ,0,21})

aAdd(aRotina5,{"Aglutinar"     ,"U_ALL_EXEC('Aglutinar')"        ,0,23})  
aAdd(aRotina5,{"Planilha"     ,"U_ALL_EXEC('Planilha AGL')"      ,0,24}) 

aAdd(aRotina,{"Importação"    ,aRotina2              ,0,09})  
aAdd(aRotina,{"Pedidos"       ,aRotina3              ,0,13}) 
aAdd(aRotina, {"Notas Fiscais",aRotina4              ,0,17}) 
If Alltrim(cPrefxTel) = 'PP'
   aAdd(aRotina, {"Financeiro"   ,aRotina5                  ,0,22})
   aAdd(aRotina,{"Composição"    ,"U_ALL_EXEC('Composição')",0,25})  
Endif
aAdd(aRotina,{"Legenda"   ,"U_ALL_LEGX()"            ,0,26})
Return (aRotina)

//------------------------------------------------------------------------
User Function ALL_LEGX()
//------------------------------------------------------------------------
Local aLegenda := {}

aAdd(aLegenda,{"BR_VERMELHO"  ,"Importação com Erro" })
aAdd(aLegenda,{"BR_VERDE"   ,"Importação com sucesso" })
aAdd(aLegenda,{"BR_AZUL"      ,"Pedido de Venda gerado" })
aAdd(aLegenda,{"BR_AZUL_CLARO","NF gerada parcial"})
aAdd(aLegenda,{"BR_LARANJA"   ,"NF gerada total"})
BrwLegenda( cCadastro, "Legenda", aLegenda )
Return Nil

//------------------------------------------------------------------------
User Function ALL_EXEC(p_Funcao)
//------------------------------------------------------------------------
Local nLinhas := Int(MLCount(PPX->PPX_ERRO,50))
 Local cTexto  := ""
 Local cCampo  := ""
 Local aPedidos := {}
  Local axPedidos := ""
 
 If Upper(Alltrim(p_Funcao)) == "EXCLUIR PV" 
    For nX:= 1 To nLinhas
        cTexto := Alltrim(MemoLine(PPX->PPX_ERRO,50,nX))
        If Alltrim(cTexto) <> ""
           If Left(cTexto,1) $ "|1|2|3|4|5|6|7|8|9|0|"
              axPedidos := "'"+Left(Alltrim(MemoLine(PPX->PPX_ERRO,50,nX)),TamSx3('C5_NUM')[1])+"',"
	          aAdd(aPedidos,Alltrim(cTexto))
	       Endif
	    Endif   
    Next nX
    If Alltrim(axPedidos) = ""
       axPedidos += "'"+PPX->PPX_PEDIDO+"','xx'"
    Endif   
    axPedidos := "("+Alltrim(axPedidos)+")"
Endif
cCampo  := ""
Do Case
      Case Upper(Alltrim(p_Funcao)) == "IMPORTAR"
           Processa({|| PPX_IMPORT()},"Importando Planilha Excel")
      Case Upper(Alltrim(p_Funcao)) == "GERAR"
           Processa({|| PPX_PVGERA()},"Gerando Pedidos de Venda")
      Case Upper(Alltrim(p_Funcao)) == "EXCLUIR PV" 
           PPX_TELAEXCL(PPX->PPX_PEDIDO,PPX->PPX_ID,axPedidos,2)
      Case Upper(Alltrim(p_Funcao)) == "EXCLUIR PV LOTE"
           PPX_TELAEXCL(Space(TamSx3('C5_NUM')[1]),PPX->PPX_ID,"",1)
      Case Upper(Alltrim(p_Funcao)) == "AVALIAR"
           Processa({|| PPX_AVALIAR()},"Analisando Pedidos de Venda com Erro")
      Case Upper(Alltrim(p_Funcao)) == "COMPOSIÇÃO"
           Processa({|| PPX_PVCOMPO()},"Montando Composição")
      Case Upper(Alltrim(p_Funcao)) == "LOTE"
           Processa({|| PPX_TELALOTE()},"Exclusão da Importação em Lote")
      Case Upper(Alltrim(p_Funcao)) == "NF"
           Processa({|| PPX_NFGERA(1,PPX->PPX_ID)},"Gerando Documento de Saída")
      Case Upper(Alltrim(p_Funcao)) == "EXCLUIR NF"
           Processa({|| PPX_NFGERA(2,PPX->PPX_ID)},"Excluindo Documento de Saída")   
      Case Upper(Alltrim(p_Funcao)) == "TRANSMITIR NF"
           Processa({|| PPX_NFGERA(3,PPX->PPX_ID)},"Transmitindo Documento de Saída")  
           dDataBase := dDBaseBKP
           cEmpAnt   := cEmpAnt
           cFilAnt   := cFilAnt
           OpenFile(cEmpAnt+cFilAnt)
           OpenSM0()
      Case Upper(Alltrim(p_Funcao)) == "HISTORICO NF"
           Processa({|| PPX_PLANNF()},"Gerando Histórico no Período")      
      Case Upper(Alltrim(p_Funcao)) == "AGLUTINAR"
           U_JbFATPPM() 
      Case Upper(Alltrim(p_Funcao)) == "PLANILHA AGL"
          Processa({|| PPX_PLANFIN(3)},"Gerando Planilha Titulos Aglutinados")                
EndCase       
Return

//------------------------------------------------------------------------
Static Function PPX_IMPORT()
//------------------------------------------------------------------------   
Local aArea     := GetArea()
Local aCampos   := {}  
Local aDados    := {}           
Local aFields  := {"PPX_PERIOD","PPX_EMPRES","PPX_EMISSA","PPX_VALOR","PPX_DESC","PPX_VENCTO","PPX_CNPJ","PPX_FANTAS","PPX_RAZAO","PPX_MASTER","PPX_LOJA","PPX_REDE","PPX_PRODUT","PPX_HIST" ,"PPX_RPS","PPX_NFSE"}
Local aCabec    := {"Periodo"   ,"Empresa"   ,"DtEmissao" ,"Valor"    ,"Desconto","Vencimento","CNPJ"    ,"Fantasia"  ,"Razão"    ,"Master"    ,"Loja"    ,"Rede"    ,"Protheus"  ,"Historico","RPS"    ,"NFS-E"}
Local aMeses    := {"JAN","FEV","MAR","ABR","MAI","JUN","JUL","AGO","SET","OUT","NOV","DEZ"}
Local astru     :={}
Local cAno      := 0
Local cAnomes   := ""
Local cArqMacro := "XLS2DBF.XLA" 
Local cPatch	:= ""
Local cDirCSV	:= GetSrvProfString("Startpath","")  
Local cEmp      := FWCodEmp()
Local cFil      := FWCodFil()
Local cId       := "" 
Local cLinha    := ""
Local cMes      := 0
Local cMsg      := ""
Local cSystem   := Upper(GetSrvProfString("RootPath",""))
Local lPPR      := .F.
Local nHandle   := 0
Local nLin      := 1 
Local nLinTit   := 1     
Local nLinTot   := 0
Local nValComp  := 0
Local nValServ  := 0
Local cNumero   := ""
Local nX        := 0    
Local nX1       := 0
Local nX2       := 0
Local nX3       := 0
Local nX4       := 0
Local nX5       := 0
Local lFirst    := .T.
Local cCNPJ     := ""
Local cPagto    := ""
Local cNatFin   := ""
Local cQuebra   := "ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ"
Local dEmissao  := CTOD("  /  /    ")
Local nValor    := 0
Local nDesconto := 0
Local nXtam     := 0

cFileC    := ""

cPathTmp    := cGetFile( '', 'Selecione Diretório onde estao os arquivos a serem processados', 0, , .F.,   GETF_LOCALHARD +  GETF_RETDIRECTORY + GETF_NETWORKDRIVE)
If Alltrim(cPathTmp) = ""
   Return
Endif
aDados  := {}
aFields := {}
aAdd(aFields,{"PPX_PERIOD","Periodo"     ,"",0,0,"UPPER(ALLTRIM(axLinha[01]))"})
aAdd(aFields,{"PPX_EMPRES","Empresa"     ,"",0,0,"LEFT(ALLTRIM(axLinha[02])+SPACE(500),TamSx3('PPX_EMPRES')[1])"})
aAdd(aFields,{"PPX_EMISSA","Dt Emissão"  ,"",0,0,"If(Valtype(axLinha[03])='D',axLinha[03],CTOD(axLinha[03]))"})
aAdd(aFields,{"PPX_VALOR" ,"Valor"       ,"",0,0,"Val(Replace(Replace(axLinha[04],'.',''),',','.'))"})
aAdd(aFields,{"PPX_DESC"  ,"Descontos"   ,"",0,0,"Val(Replace(Replace(axLinha[05],'.',''),',','.'))"})
aAdd(aFields,{"PPX_VENCTO","Dt Vencto"   ,"",0,0,"If(Valtype(axLinha[06])='D',axLinha[06],CTOD(axLinha[06]))"})
aAdd(aFields,{"PPX_CNPJ"  ,"CNPJ"        ,"",0,0,"REPLACE(REPLACE(REPLACE(axLinha[07],'.',''),'-',''),'/','')"})
aAdd(aFields,{"PPX_FANTAS","Fantasia"    ,"",0,0,"LEFT(ALLTRIM(axLinha[08])+SPACE(500),TamSx3('PPX_FANTAS')[1])"})
aAdd(aFields,{"PPX_RAZAO" ,"Razão Social","",0,0,"LEFT(ALLTRIM(axLinha[09])+SPACE(500),TamSx3('PPX_RAZAO')[1])"})
aAdd(aFields,{"PPX_MASTER","Master"      ,"",0,0,"LEFT(ALLTRIM(axLinha[10])+SPACE(500),TamSx3('PPX_MASTER')[1])"})
aAdd(aFields,{"PPX_LOJA"  ,"Loja"        ,"",0,0,"STRZERO(VAL(axLinha[11]),TamSx3('PPX_LOJA')[1])"})
aAdd(aFields,{"PPX_REDE"     ,"Rede"     ,"",0,0,"LEFT(ALLTRIM(axLinha[12])+SPACE(500),TamSx3('PPX_REDE')[1])"})
aAdd(aFields,{"PPX_PRODUT","Produto"     ,"",0,0,"LEFT(ALLTRIM(axLinha[13])+SPACE(500),TamSx3('PPX_PRODUT')[1])"})
aAdd(aFields,{"PPX_HIST"  ,"Historico"   ,"",0,0,"If(Len(axLinha)>13,LEFT(ALLTRIM(axLinha[14])+SPACE(500),TamSx3('PPX_HIST')[1]),'')"})
aAdd(aFields,{"PPX_RPS"   ,"RPS"         ,"",0,0,"If(Len(axLinha)>13,STRZERO(VAL(axLinha[15]),TamSx3('PPX_RPS')[1]),'')"})
aAdd(aFields,{"PPX_NFSE"  ,"NFS-E"       ,"",0,0,"If(Len(axLinha)>13,STRZERO(VAL(axLinha[16]),TamSx3('PPX_NFSE')[1]),'')"})

dbSelectArea("SX3")
SX3->(DbSetOrder(2))
For nX1 := 1 to Len(aFields)
    If SX3->(DbSeek(aFields[nX1,1]))
       aFields[nX1,3] := SX3->X3_TIPO
       aFields[nX1,4] := SX3->X3_TAMANHO
       aFields[nX1,5] := SX3->X3_DECIMAL
    Endif
Next nX
aArqPPM := {}
aArq := directory(Alltrim(cPathTmp)+"*.csv") 
If Len(aArq) < 1
   Alert("Não foram encontrados arquivos, nem Extensão [CSV]")
   Return
Endif   
For nX1 := 1 to Len(aArq) 
    If Upper(Left(Alltrim(aArq[nX1,1]),2)) $ "PP|PZ|"
      cFileC := Alltrim(cPathTmp)+Replace(Alltrim(aArq[nX1,1]),"\\","\")
       aAdd(aArqPPM,aArq[nX1])
       Loop
    Endif   
    cFileC := Alltrim(cPathTmp)+Replace(Alltrim(aArq[nX1,1]),"\\","\")
    nHandle  := Ft_Fuse(cFileC)
    If nHandle == -1
       Help(,,"Help","Importação de Pedidos de Venda", "Arquivo não existe ou está em uso["+Alltrim(aArq[nX1,1])+"]", 1, 0)
       Return 
    Endif    
    aTitulo := U_ETX_EMPRE(Left(aArq[nX1,1],2))
    If Alltrim(aTitulo[1]) = ""
       Alert("Empresa não cadastrada na tabela SX5==>"+Alltrim(aTitulo[1])+"-"+Alltrim(aTitulo[2])+"-"+Alltrim(aTitulo[3])+"-"+Alltrim(aTitulo[4])+"-"+Alltrim(aTitulo[5]))
       Return
    Endif 
    dbCloseAll()
    cPrefxArq := aTitulo[1]
    cEmpX     := aTitulo[2]
    cFilX     := aTitulo[3]
    cEmpAnt   := cEmpX
    cFilAnt   := cFilX
    OpenFile(cEmpAnt+cFilAnt)
    OpenSM0()
    dDataBase := Date()
    nLin    := 0
    nLinTit := 1
    nLinTot := 0
    axLinha  := {}
    cLinha  := ""
    lFirst  := .T.  
	cQuebra := "ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ"
	aDados := {}
	Ft_FGoTop()                                                         
	nLinTot := FT_FLastRec()-1
	ProcRegua(nLinTot)
	While nLinTit > 0 .AND. !Ft_FEof() //Pula as linhas de cabeçalho
	      cLinha := Ft_FReadLn()
	      cLinha  := '{"'+Alltrim(cLinha)+'"}'
	      cLinha  := StrTran(cLinha,';','","')//adiciona o cLinha no array trocando o delimitador ; por , para ser reconhecido como elementos de um array 
	      aCabec := aClone(&(cLinha))
	      nLinTit--
	      Ft_FSkip()
	EndDo
	While nLinTot > 0 .AND. !Ft_FEof() //percorre todas linhas do arquivo csv
	      IncProc("Carregando Linha "+Alltrim(aArq[nX1,1])+" - "+AllTrim(Str(nLin))+" de "+AllTrim(Str(nLinTot)))
	      cLinha := Ft_FReadLn()
	      If Empty(AllTrim(StrTran(cLinha,';','')))
	         Ft_FSkip()
	         Loop
	      EndIf
	      axLinha  := {}
	      nLin++
	      cLinha  := '{"'+Alltrim(cLinha)+'"}'
	      cLinha  := StrTran(cLinha,';','","')//adiciona o cLinha no array trocando o delimitador ; por , para ser reconhecido como elementos de um array 
	      axLinha := aClone(&(cLinha)) 
	      nXtam   := Len(axLinha)
	      If nXtam > Len(aFields)
	         nXtam := Len(aFields)
	      Endif 
	      If Alltrim(axLinha[1]) = ""
	         Ft_FSkip()
	         Loop
	      EndIf 
	      For nX3 := 1 to nXtam
	          axLinha[nX3]:= &(aFields[nX3,6]) 
	          nX3 := nX3
	      Next nX3
	      If ValType(axLinha[1]) =="C"
	         If Len(Alltrim(axLinha[1])) = 10
	            cLinha := DTOS(CTOD(axLinha[1]))
	            axLinha[1] := Substr(cLinha,5,2)+"/"+Substr(cLinha,1,4)
	         Endif 	  
	      Endif       
	      If ValType(axLinha[1]) =="D"
	         cLinha := DTOS(axLinha[1])
	         axLinha[1] := Substr(cLinha,5,2)+"/"+Substr(cLinha,1,4)
	      Endif   	         
	      
	      nX2 := aScan(aMeses,Upper(Left(axLinha[1],3)))
	      If nX2 > 0
	         axLinha[1] := STRZERO(nX2,2)+"/"+"20"+Substr(axLinha[1],5,4)
	      Endif  
	      For nX3 := nXtam to 17
	          aAdd(axLinha,"")
	      Next nX3    
	      cPagto      := POSICIONE("SB1",1,xFilial("SB1")+axLinha[13],"B1_XCOND") 
          cNatFin     := POSICIONE("SB1",1,xFilial("SB1")+axLinha[13],"B1_XNAT")
	      axLinha[17] := Alltrim(axLinha[1])+Alltrim(axLinha[2])+Alltrim(axLinha[10])+Alltrim(axLinha[11])+Alltrim(axLinha[12])+Alltrim(axLinha[7])+DTOS(axLinha[3])+DTOS(axLinha[6])+Alltrim(cPagto)+Alltrim(cNatFin)
	      aAdd(aDados,axLinha)
	      FT_FSkip()
	EndDo 
	nLin  := 0
	nXtam := aDados[1,Len(aDados[1])]
	aSort(aDados,,,{|x,y| x[17] < y[17]})
	For nX2 := 1 to LEN(aDados)
	    If Alltrim(aDados[nX2,17]) <> Alltrim(cQuebra)
	       nLin     := 0
          For nLin := 1 to 10000
              cNumero := GetSxeNum('PPX','PPX_ID')
              ConfirmSX8()
              dbSelectArea("PPX")
              PPX->(dbSetOrder(1))
              If !(PPX->(dbSeek(xFilial("PPX")+cNumero)))
                  nLin := 10000
              Endif
           Next nLin 
           nLin := 0
	       cMsg := ""
	       dbSelectArea("PPX") 
	       Reclock("PPX",.T.)
	       For nX3 := 1 to LEN(aFields)
	           PPX->&(aFields[nX3,1]) := aDados[nX2,nX3]
	       Next nX2   
	       
	       PPX->PPX_FILIAL := XFILIAL("PPX")
	       PPX->PPX_ID     := cNumero
	       PPX->PPX_STATUS := "1"
	       PPX->PPX_AGLUT:= "N"
	       PPX->PPX_PREFIX := cPrefxArq
	       PPX->PPX_DATA   := dDataBase
	       PPX->PPX_FILE   := Alltrim(cDirCSV)+Alltrim(aArq[nX1,1])
	       PPX->PPX_LINE   := STRZERO(1,TamSx3("PPX_LINE")[1])
	       PPX->PPX_CODNAT := POSICIONE("SB1",1,xFilial("SB1")+PPX->PPX_PRODUT,"B1_XNAT")
	       PPX->PPX_CONDPG := POSICIONE("SB1",1,xFilial("SB1")+PPX->PPX_PRODUT,"B1_XCOND")
           cMsg            := U_ALLVERIF("PPX")
           PPX->PPX_VALOR  := 0
	       PPX->PPX_DESC   := 0
           If Alltrim(cMsg) <> ""
              PPX->PPX_ERRO   := Alltrim(cMsg)
              PPX->PPX_STATUS := "0"
           Endif    
	       PPX->(MsUnlock())
	    Endif       
	    axLinha   := aDados[nX2]           
	    cPagto    := POSICIONE("SB1",1,xFilial("SB1")+aDados[nX2,13],"B1_XCOND")
        cNatFin   := POSICIONE("SB1",1,xFilial("SB1")+aDados[nX2,13],"B1_XNAT")
        cQuebra   := Alltrim(axLinha[1])+Alltrim(axLinha[2])+Alltrim(axLinha[10])+Alltrim(axLinha[11])+Alltrim(axLinha[12])+Alltrim(axLinha[7])+DTOS(axLinha[3])+DTOS(axLinha[6])+Alltrim(cPagto)+Alltrim(cNatFin)
	    nLin++
	    cMsg := ""
	    dbSelectArea("PPZ") 
	    Reclock("PPZ",.T.)
        PPZ->PPZ_FILIAL := XFILIAL("PPZ")
        PPZ->PPZ_ID     := cNumero
        PPZ->PPZ_EMPRES := PPX->PPX_EMPRES
        PPZ->PPZ_ITEM   := STRZERO(nLin,TamSx3("PPZ_ITEM")[1])
        PPZ->PPZ_CODPRO := aDados[nX2,13]
        PPZ->PPZ_QTDVEN := 1
        PPZ->PPZ_PRCVEN := aDados[nX2,4]
        PPZ->PPZ_PRUNIT := aDados[nX2,4]
        PPZ->PPZ_TPOP   := "F"
        PPZ->PPZ_SUGENT := dDataBase
        PPZ->PPZ_VALOR  := aDados[nX2,04]
        PPZ->PPZ_VLDESC := aDados[nX2,05]
        PPZ->PPZ_LJREDE := aDados[nX2,11]
        PPZ->PPZ_EMISSA := aDados[nX2,03]
        PPZ->PPZ_VENCTO := aDados[nX2,06]
        PPZ->PPZ_CNPJ   := aDados[nX2,07]
        PPZ->PPZ_RAZAO  := aDados[nX2,08]
        PPZ->PPZ_FANTA  := aDados[nX2,09]
        PPZ->PPZ_MASTER := aDados[nX2,10]
        PPZ->PPZ_REDE   := aDados[nX2,12]
        PPZ->PPZ_TES    := POSICIONE("SB1",1,xFilial("SB1")+aDados[nX2,13],"B1_TS")
        PPZ->PPZ_LOCAL  := POSICIONE("SB1",1,xFilial("SB1")+aDados[nX2,13],"B1_LOCPAD")
        PPZ->PPZ_CLI    := POSICIONE("SA1",3,xFilial("SA1")+aDados[nX2,07],"A1_COD")
        PPZ->PPZ_LOJA   := POSICIONE("SA1",3,xFilial("SA1")+aDados[nX2,07],"A1_LOJA")
        PPZ->PPZ_PEDIDO := ""
        PPZ->(MsUnlock())
        dbSelectArea("PPX")
        PPX->(dbSetOrder(1))
        PPX->(dbSeek(xFilial("PPX")+cNumero))
        RecLock("PPX",.F.)
        PPX->PPX_VALOR  := PPX->PPX_VALOR + aDados[nX2,4]
	    PPX->PPX_DESC   := PPX->PPX_DESC  + aDados[nX2,5]
        PPX->(MsUnlock())
	Next nX2    
	FT_FUse() 
	WFMoveFiles( cPathTmp+aArq[nX1,1],cPathTmp+"Importados") 
	Ferase(cPathTmp+aArq[nX1,1])
Next nX1  
If Len(aArqPPM) > 0
   dbCloseAll()
   aTitulo := U_ETX_EMPRE("PP")
   If Alltrim(aTitulo[1]) = ""
      Alert("Empresa não cadastrada na tabela SX5==>"+Alltrim(aTitulo[1])+"-"+Alltrim(aTitulo[2])+"-"+Alltrim(aTitulo[3])+"-"+Alltrim(aTitulo[4])+"-"+Alltrim(aTitulo[5]))
      Return
   Endif  
    cPrefxArq := aTitulo[1]
   cEmpX      := aTitulo[2]
   cFilX      := aTitulo[3]
   cEmpAnt    := cEmpX
   cFilAnt    := cFilX
   OpenFile(cEmpAnt+cFilAnt)
   OpenSM0()
   
   PPX_IMPPPM(cPathTmp,aArqPPM)
Endif  
RestArea( aArea )
dbCloseAll()
cEmpAnt := cEmpBKP   
cFilAnt := cFilBKP 
OpenFile(cEmpAnt+cFilAnt)
OpenSM0() 
return 

//-------------------------------------------------------------------------
Static Function PPX_IMPPPM(cPathTmp,aArqPPM)
//-------------------------------------------------------------------------
Local aArea     := GetArea()
Local aCampos   := {}  
Local aDados    := {}           
Local aFields  := {"PPX_PERIOD","PPX_EMPRES","PPX_EMISSA","PPX_VALOR","PPX_DESC","PPX_VENCTO","PPX_CNPJ","PPX_FANTAS","PPX_RAZAO","PPX_MASTER","PPX_LOJA","PPX_REDE","PPX_PRODUT","PPX_TOTFAT" ,"PPX_CNPJAJS","PPX_AGLUT"}
Local aCabec    := {"Periodo"   ,"Empresa"   ,"DtEmissao" ,"Valor"    ,"Desconto","Vencimento","CNPJ"    ,"Fantasia"  ,"Razão"    ,"Master"    ,"Loja"    ,"Rede"    ,"Protheus"  ,"Vr Tot Fat","CNPJ Raiz"  ,"Aglutina"}
Local aMeses    := {"JAN","FEV","MAR","ABR","MAI","JUN","JUL","AGO","SET","OUT","NOV","DEZ"}
Local astru     :={}
Local cAno      := 0
Local cAnomes   := ""
Local cArqMacro := "XLS2DBF.XLA" 
Local cPatch	:= ""
Local cDirCSV	:= GetSrvProfString("Startpath","")  
Local cEmp      := FWCodEmp()
Local cFil      := FWCodFil()
Local cId       := "" 
Local cLinha    := ""
Local cMes      := 0
Local cMsg      := ""
Local cSystem   := Upper(GetSrvProfString("RootPath",""))
Local lPPR      := .F.
Local nHandle   := 0
Local nLin      := 1 
Local nLinTit   := 1     
Local nLinTot   := 0
Local nTotFat   := 0
Local nFatAjs   := 0
Local nValComp  := 0
Local nValServ  := 0
Local nValor    := 0
Local nTxADM    := 0
Local nTxSrv    := 0
Local nTxMid    := 0
Local nDesconto := 0
Local cPeriodo  := ""
Local cNumero   := ""
Local nX        := 0    
Local nX1       := 0
Local nX2       := 0
Local nX3       := 0
Local nX4       := 0
Local nX5       := 0
Local lFirst    := .T.
Local cCNPJ     := ""
Local cPagto    := ""
Local cNatFin   := ""
Local cQuebra   := "ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ"
Local dEmissao  := CTOD("  /  /    ")
Local nXtam     := 0
Local cPPMNum   := ""

cFileC    := ""
    aDados  := {} 
aFields     := {}
aAdd(aFields,{"PPX_PERIOD","Periodo"     ,"",0,0,"UPPER(ALLTRIM(axLinha[01]))"})
aAdd(aFields,{"PPX_EMPRES","Empresa"     ,"",0,0,"LEFT(ALLTRIM(axLinha[02])+SPACE(500),TamSx3('PPX_EMPRES')[1])"})
aAdd(aFields,{"PPX_EMISSA","Dt Emissão"  ,"",0,0,"If(Valtype(axLinha[03])='D',axLinha[03],CTOD(axLinha[03]))"})
aAdd(aFields,{"PPX_VALOR" ,"Valor"       ,"",0,0,"Val(Replace(Replace(axLinha[04],'.',''),',','.'))"})
aAdd(aFields,{"PPX_DESC"  ,"Descontos"   ,"",0,0,"Val(Replace(Replace(axLinha[05],'.',''),',','.'))"})
aAdd(aFields,{"PPX_VENCTO","Dt Vencto"   ,"",0,0,"If(Valtype(axLinha[06])='D',axLinha[06],CTOD(axLinha[06]))"})
aAdd(aFields,{"PPX_CNPJ"  ,"CNPJ"        ,"",0,0,"REPLACE(REPLACE(REPLACE(axLinha[07],'.',''),'-',''),'/','')"})
aAdd(aFields,{"PPX_FANTAS","Fantasia"    ,"",0,0,"LEFT(ALLTRIM(axLinha[08])+SPACE(500),TamSx3('PPX_FANTAS')[1])"})
aAdd(aFields,{"PPX_RAZAO" ,"Razão Social","",0,0,"LEFT(ALLTRIM(axLinha[09])+SPACE(500),TamSx3('PPX_RAZAO')[1])"})
aAdd(aFields,{"PPX_MASTER","Master"      ,"",0,0,"LEFT(ALLTRIM(axLinha[10])+SPACE(500),TamSx3('PPX_MASTER')[1])"})
aAdd(aFields,{"PPX_LOJA"  ,"Loja"        ,"",0,0,"STRZERO(VAL(axLinha[11]),TamSx3('PPX_LOJA')[1])"})
aAdd(aFields,{"PPX_REDE"  ,"Rede"        ,"",0,0,"LEFT(ALLTRIM(axLinha[12])+SPACE(500),TamSx3('PPX_REDE')[1])"})
aAdd(aFields,{"PPX_PRODUT","Produto"     ,"",0,0,"LEFT(ALLTRIM(axLinha[13])+SPACE(500),TamSx3('PPX_PRODUT')[1])"})
aAdd(aFields,{"PPX_TOTFAT","Vr Tot Fat"  ,"",0,0,"If(Len(axLinha)>15,Val(Replace(Replace(axLinha[14],'.',''),',','.')),0)"})
aAdd(aFields,{"PPX_CNPJAG","CNPJ Raiz"   ,"",0,0,"If(Len(axLinha)>15,REPLACE(REPLACE(REPLACE(axLinha[15],'.',''),'-',''),'/',''),'0')"})
aAdd(aFields,{"PPX_AGLUT" ,"Aglut"       ,"",0,0,"If(Len(axLinha)>15,Alltrim(axLinha[16]),'N')"})
aAdd(aFields,{"PPX_FILE"  ,"Nome arquivo","",0,0,"If(Len(axLinha)>15,Alltrim(cFileC),Alltrim(cFileC))"})
aAdd(aFields,{"PPX_LINE"  ,"Linha"       ,"",0,0,"If(Len(axLinha)>15,STRZERO(1,TamSx3('PPX_LINE')[1]),STRZERO(1,TamSx3('PPX_LINE')[1]))"})
          
cPeriodo := "99/9999"
nTotFat   := 0
nFatAjs   := 0
nValComp  := 0
nValServ  := 0
dbSelectArea("SX3")
SX3->(DbSetOrder(2))
For nX1 := 1 to Len(aFields)
    If SX3->(DbSeek(aFields[nX1,1]))
       aFields[nX1,3] := SX3->X3_TIPO
       aFields[nX1,4] := SX3->X3_TAMANHO
       aFields[nX1,5] := SX3->X3_DECIMAL
    Endif
Next nX
aArq := aClone(aArqPPM)
If Len(aArq) < 1
   Alert("Não foram encontrados arquivos com Prefixo [PP/PX], nem Extensão [CSV]")
   Return
Endif   
For nX1 := 1 to Len(aArq)  
    cFileC := Alltrim(cPathTmp)+Replace(Alltrim(aArq[nX1,1]),"\\","\")
    nHandle  := Ft_Fuse(cFileC)
    If nHandle == -1
       Help(,,"Help","Importação de Pedidos de Venda-PPM", "Arquivo não existe ou está em uso["+Alltrim(aArq[nX1,1])+"]", 1, 0)
       Return 
    Endif    
    dDataBase := Date()
    nLin    := 0
    nLinTit := 1
    nLinTot := 0 
    axLinha := {}
    cLinha  := ""
    lFirst  := .T.  
	cQuebra := "ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ"
	Ft_FGoTop()                                                         
	nLinTot := FT_FLastRec()-1
	ProcRegua(nLinTot)
	While nLinTit > 0 .AND. !Ft_FEof() //Pula as linhas de cabeçalho
	      cLinha := Ft_FReadLn()
	      cLinha  := '{"'+Alltrim(cLinha)+'"}'
	      cLinha  := StrTran(cLinha,';','","')//adiciona o cLinha no array trocando o delimitador ; por , para ser reconhecido como elementos de um array 
	      aCabec := aClone(&(cLinha))
	      nLinTit--
	      Ft_FSkip()
	EndDo
	While nLinTot > 0 .AND. !Ft_FEof() //percorre todas linhas do arquivo csv
	      IncProc("Carregando Linha "+Alltrim(aArq[nX1,1])+" - "+AllTrim(Str(nLin))+" de "+AllTrim(Str(nLinTot)))
	      cLinha := Ft_FReadLn()
	      If Empty(AllTrim(StrTran(cLinha,';','')))
	         Ft_FSkip()
	         Loop
	      EndIf
	      nLin++
	      cLinha  := '{"'+Alltrim(cLinha)+'"}'
	      cLinha  := StrTran(cLinha,';','","')//adiciona o cLinha no array trocando o delimitador ; por , para ser reconhecido como elementos de um array 
	      axLinha := aClone(&(cLinha)) 
	      If Alltrim(axLinha[1]) = ""
	         Ft_FSkip()
	         Loop
	      Endif   
	      For nX3 := Len(axLinha) to Len(aFields)
	         aAdd(axLinha,"")
	      Next nX3
	      For nX3 := 1 to Len(aFields)  
	          axLinha[nX3]:= &(aFields[nX3,6])  //aFields[nX3,1]
	           nX3 := nX3
	      Next nX3 
	      If Left(aArq[nX1,1],2) = "PZ"
	         axLinha[16] := "I"
	      Endif   
	      nX2 := aScan(aMeses,Upper(Left(axLinha[1],3)))
	      If nX2 > 0
	         axLinha[1] := STRZERO(nX2,2)+"/"+"20"+Substr(axLinha[1],5,4)
	      Endif   
	      nTotal   := axLinha[14]
	      If Alltrim(fLerSX5("Z3",axLinha[12])) = "CH" 
	         axLinha[16] := "C"
	      Else   
	         If axLinha[16] <> "I"
	            nFatAjs += axLinha[4] - axLinha[5]
	         Endif   
	      Endif 
	      If Alltrim(axLinha[16]) = "N"  
	         axLinha[16] := "G"  
	      Endif
	      If Alltrim(axLinha[16]) $ "C|G|S"
	         cPeriodo := Substr(axLinha[1],4,4)+Substr(axLinha[1],1,2) 
	      Endif 
	      aAdd(aDados,axLinha) 
	      FT_FSkip()
	EndDo    
	FT_FUse() 
Next nX1  
If Alltrim(cPeriodo) <> "99/9999"
   dbSelectArea("PAQ")
   PAQ->(dbSetOrder(1))
   If !(PAQ->(dbSeek(xFilial("PAQ")+cPeriodo)))
      MsgStop ("Falta cadastrar os Valores de Competência para o Período [ "+cPeriodo+"]","Erro")
      Return
   Endif  
   nValComp := PAQ->PAQ_VALOR 
Endif
nValor   := 0
nDesc    := 0
nTxADM   := 0
nTxSrv   := 0
nTxMid   := 0
cPPMNum := GetSxeNum('PPX','PPX_ID') 
ConfirmSX8()
dbSelectArea("PPX")
RecLock("PPX",.T.)
PPX->PPX_FILIAL  :=xFilial("PPX")
PPX->PPX_ID      := cPPMNum
PPX->PPX_LINE    := STRZERO(1,TamSx3('PPX_LINE')[1])
PPX->PPX_PREFIX  := "PP"
PPX->PPX_STATUS  := "1"
PPX->PPX_AGLUT   := "S"
PPX->PPX_DATA    := dDataBase
PPX->PPX_VALOR   := 0
PPX->PPX_DESC    := 0
PPX->PPX_TOTFAT  := 0
PPX->PPX_TOTAJS  := nFatAjs
PPX->PPX_VRCOMP  := nValComp
PPX->(MsUnlock()) 
nLin := 0
For nX1 := 1 to Len(aDados)
    If aDados[nX1,16] $ "C|G|I"
	   cNumero := GetSxeNum('PPX','PPX_ID')
	   ConfirmSX8()
	   nTxADM := 0
	   nTxSRV := 0
	   nTxMid := 0
	Else
	   cNumero := cPPMNum  
      
	Endif   
	cMsg := ""
	dbSelectArea("PPX") 
	PPX->(dbSetOrder(1))
	If PPX->(dbSeek(xFilial("PPX")+cNumero))
	   Reclock("PPX",.F.)
	Else   
	   Reclock("PPX",.T.)
	Endif   
	For nX2 := 1 to LEN(aFields)
	    PPX->&(aFields[nX2,1]) := aDados[nX1,nX2]
	Next nX2  
	cMsg    := U_ALLVERIF("PPX")
	PPX->PPX_FILIAL := XFILIAL("PPX")
	PPX->PPX_ID     := cNumero
	PPX->PPX_STATUS := "1"
	PPX->PPX_ERRO   := ""
	PPX->PPX_CNPJ   := If(PPX->PPX_AGLUT="S",aDados[nX1,15],aDados[nX1,07])
	PPX->PPX_DATA   := dDataBase
	PPX->PPX_TOTAJS := If(aDados[nX1,16]="C",0,nFatAjs)
	PPX->PPX_VRCOMP := If(aDados[nX1,16]="C",0,nValComp)
	PPX->PPX_PREFIX := "PP"
	PPX->PPX_CODNAT := POSICIONE("SB1",1,xFilial("SB1")+PPX->PPX_PRODUT,"B1_XNAT")
	PPX->PPX_CONDPG := POSICIONE("SB1",1,xFilial("SB1")+PPX->PPX_PRODUT,"B1_XCOND")
	PPX->PPX_AGLUT  := aDados[nX1,16]
    cMsg            := U_ALLVERIF("PPX")
    If Alltrim(cMsg) <> ""
       PPX->PPX_ERRO   := Alltrim(cMsg)
       PPX->PPX_STATUS := "0"
    Endif 
    dbSelectArea("PPZ") 
	Reclock("PPZ",.T.)
    PPZ->PPZ_FILIAL := XFILIAL("PPZ")
    PPZ->PPZ_ID     := cNumero
    PPZ->PPZ_EMPRES := PPX->PPX_EMPRES
    PPZ->PPZ_ITEM   := STRZERO(1,TamSx3("PPZ_ITEM")[1])
    PPZ->PPZ_CODPRO := PPX->PPX_PRODUT
    PPZ->PPZ_QTDVEN := 1
    PPZ->PPZ_PRCVEN := aDados[nX1,4]- aDados[nX1,5]
    PPZ->PPZ_PRUNIT := aDados[nX1,4]
    PPZ->PPZ_TPOP   := "F"
    PPZ->PPZ_SUGENT := dDataBase
    PPZ->PPZ_LJREDE := aDados[nX1,11]
    PPZ->PPZ_EMISSA := aDados[nX1,03]
    PPZ->PPZ_VENCTO := aDados[nX1,06]
    PPZ->PPZ_CNPJ   := aDados[nX1,07]
    PPZ->PPZ_RAZAO  := aDados[nX1,08]
    PPZ->PPZ_FANTA  := aDados[nX1,09]
    PPZ->PPZ_MASTER := aDados[nX1,10]
    PPZ->PPZ_REDE   := aDados[nX1,12]
    PPZ->PPZ_VALOR  := aDados[nX1,04]
    PPZ->PPZ_VLDESC := aDados[nX1,05]
    PPZ->PPZ_TES    := POSICIONE("SB1",1,xFilial("SB1")+PPX->PPX_PRODUT,"B1_TS")
    PPZ->PPZ_LOCAL  := POSICIONE("SB1",1,xFilial("SB1")+PPX->PPX_PRODUT,"B1_LOCPAD")
    PPZ->PPZ_CLI    := POSICIONE("SA1",3,xFilial("SA1")+aDados[nX1,7],"A1_COD")
    PPZ->PPZ_LOJA   := POSICIONE("SA1",3,xFilial("SA1")+aDados[nX1,7],"A1_LOJA")
    If aDados[nX1,16] $ "G|S|"
       If nFatAjs > 0
          PPZ->PPZ_TXADM := (PPZ->PPZ_VALOR - PPZ->PPZ_VLDESC) * 10 / 100
          PPZ->PPZ_TXSRV := (PPZ->PPZ_VALOR - PPZ->PPZ_VLDESC) / nFatAjs * nValComp
          PPZ->PPZ_TXMID := (PPZ->PPZ_VALOR - PPZ->PPZ_VLDESC) - PPZ->PPZ_TXADM - PPZ->PPZ_TXSRV
          nValor  += PPZ->PPZ_VALOR 
          nDesc   += PPZ->PPZ_VLDESC
          nTxADM  += PPZ->PPZ_TXADM
          nTxSrv  += PPZ->PPZ_TXSRV
          nTxMid  += PPZ->PPZ_TXMID
          nTotAdm += PPZ->PPZ_TXADM
          nTotSrv += PPZ->PPZ_TXSRV
         nTotMid  += PPZ->PPZ_TXMID
       Endif    
    Endif   
    PPZ->PPZ_PEDIDO := ""
    If aDados[nX1,16] $ "G|S"
       nLin ++
       PPZ->PPZ_ITEM  := STRZERO(nLin,TamSx3("PPZ_ITEM")[1])
       PPX->PPX_VALOR := Round(nValor,2)
       PPX->PPX_DESC  := Round(nDesc ,2)
       PPX->PPX_TXADM := Round(nTxADM,2)
       PPX->PPX_TXSRV := Round(nTxSRV,2)
       PPX->PPX_TXMID := Round(nTxMid,2)
       PPX->PPX_LINE  := STRZERO(1,TamSx3('PPX_LINE')[1])
    Endif  
    PPZ->(MsUnlock())  
    PPX->(MsUnlock())  
    If aDados[nX1,16] $ "C|G|I|"  
       nValor     := 0
       nDesc      := 0
        nTxADM    := 0
        nTxSrv    := 0
        nTxMid    := 0
        nLin      := 0
    Endif    
Next nX1  
If Alltrim(cPPMNum) <> ""
   dbSelectArea("PPX") 
	PPX->(dbSetOrder(1))
	If PPX->(dbSeek(xFilial("PPX")+cPPMNum))
	   Reclock("PPX",.F.)
	   nTxSrv := nValComp - nTotSrv
	   PPX->PPX_TXSRV := PPX->PPX_TXSRV  + nTxSrv
	   PPX->PPX_TXMID := PPX->PPX_VALOR - PPX->PPX_DESC -PPX->PPX_TXADM - PPX->PPX_TXSRV 
       dbSelectArea("PPZ") 
	   PPZ->(dbSetOrder(1))
	   If PPZ->(dbSeek(xFilial("PPZ")+cPPMNum))	  
	       Reclock("PPZ",.F.) 
	       PPZ->PPZ_TXSRV := PPZ->PPZ_TXSRV + nTxSrv
	       PPZ->PPZ_TXMID := PPZ->PPZ_VALOR - PPZ->PPZ_VLDESC -PPZ->PPZ_TXADM - PPZ->PPZ_TXSRV 
	       PPZ->(MsUnlock())   
	   Endif
	   PPX->(MsUnlock()) 
    Endif
Endif
If Len(aArq) > 0
   For nX1 := 1 to Len(aArq)
       cFileC := Alltrim(cPathTmp)+Alltrim(aArq[nX1,1])
       WFMoveFiles(cFileC,cPathTmp+"Importados") 
       nX2 := nX2
       Ferase(cFileC)
   Next nX1 
   For nX1 := 1 to Len(aArq)
       cFileC := Alltrim(cPathTmp)+Alltrim(aArq[nX1,1])
       WFMoveFiles(cFileC,cPathTmp+"Importados") 
       nX2 := nX2
       Ferase(cFileC)
   Next nX1 
Endif  
return 

//------------------------------------------------------------------------
Static Function PPX_AVALIAR()
//------------------------------------------------------------------------  
Local cQryAux   := ""
Local cPagto    := ""
Local cNatFin   := ""
Local cMsg      := ""
Local nLin      := 0
Local nTotal    := 0

If Select("QRY_AUX")>0
   DbSelectArea("QRY_AUX")
   DbCloseArea()
Endif
cQryAux += " SELECT * FROM "+RetSqlName("PPX")+ Chr(13)+Chr(10)
cQryAux += " WHERE D_E_L_E_T_ = ''"	+ Chr(13)+Chr(10)
cQryAux += " AND PPX_FILIAL = '"+xFilial("PPX")+"'"	+ Chr(13)+Chr(10)
cQryAux += " AND PPX_STATUS IN ('0','1')"+ Chr(13)+Chr(10)
cQryAux += " ORDER BY PPX_ID"	+ Chr(13)+Chr(10)
cQryAux := ChangeQuery(cQryAux)
TCQuery cQryAux New Alias "QRY_AUX"
dbselectarea("QRY_AUX")                   
Count to nCount 
QRY_AUX->(DbGoTop())
ProcRegua(nCount)
While QRY_AUX->(!EOF())  
      cMsg := ""
      nLin += 1
      IncProc("Avaliando: "+Alltrim(STR(nLin,5))+" de "+Alltrim(STR(nTotal))+"-"+QRY_AUX->PPX_ID+"-"+Alltrim(QRY_AUX->PPX_FANTAS))
      dbSelectArea("PPX")
      PPX->(dbSetOrder(1))
      PPX->(dbSeek(QRY_AUX->PPX_FILIAL+QRY_AUX->PPX_ID))
      RecLock("PPX",.F.)
      PPX->PPX_STATUS := "1"
      PPX->PPX_ERRO   := ""
      PPX->PPX_CONDPG := SB1->B1_XCOND
      PPX->PPX_CODNAT := SB1->B1_XNAT 
      PPX->PPX_ERRO   := ""
      cMsg            := ""
      cMsg            := U_ALLVERIF("PPX")
      If Alltrim(cMsg) <> ""
         PPX->PPX_ERRO   := cMsg
         PPX->PPX_STATUS := "0"
      Else  
         dbSelectArea("PPZ")
         PPZ->(dbSetorder(1))
         PPZ->(dbSeek(PPX->PPX_FILIAL+PPX->PPX_ID))
         While !PPZ->(Eof()) .AND. (PPZ->PPZ_FILIAL+PPZ->PPZ_ID == PPX->PPX_FILIAL+PPX->PPX_ID)
               RecLock("PPZ",.F.)
               PPZ->PPZ_EMPRES := PPX->PPX_EMPRES
               PPZ->PPZ_CLI    := SA1->A1_COD
               PPZ->PPZ_LOJA   := SA1->A1_LOJA
               PPZ->PPZ_TES    := SB1->B1_TS
               PPZ->PPZ_LOCAL  := SB1->B1_LOCPAD  
               If PPZ->PPZ_VLDESC > PPZ->PPZ_VALOR
                  PPX->PPX_ERRO += If(Alltrim(PPX->PPX_ERRO)<>""," / ","")+"Descontos"
               Endif                
               PPZ->(MsUnlock()) 
               PPZ->(DbSkip())      
         EndDo
      Endif    
      If Alltrim(PPX->PPX_ERRO) <> ""
         PPX->PPX_STATUS := "0"
      Endif  
      PPX->(MsUnlock())                 
     QRY_AUX->(DbSkip())      
EndDo
Return 

//-------------------------------------------------------------------------
Static Function PPX_PVGERA()    //gerando Pedido de Venda
//-------------------------------------------------------------------------
Local nX1        := 0
Local nX2        := 0
Local nItem      := 0
Local cNumPed    := ""
Local axLinha    := {}
Local aItens     := {}
Local aCabec     := {}
Local aPPX       := {}
Local nXX        := 0
Local nYY        := 0
Local nZZ        := 0
Local cTexto     := 0
Local cMarca     := oMark:Mark()
Local lInverte   := oMark:IsInvert()
Local nCount     := 0
Local aCNPJ      := {}
Local nRecno     := 0

cLog		 := ''
cArqLog    := 'LOGCSV.LOG'
dbSelectArea("PPX")
PPX->(dbGoTop())
dbSelectArea("PPX")
While !PPX->(EOF())
       cPrefxTel := U_ETX_PREFIXO()
       If Alltrim(PPX->PPX_PREFIX) = Alltrim(cPrefxTel)
          If oMark:IsMark(cMarca) .AND. PPX->PPX_STATUS $ "|1|3|"
             nCount++
             aAdd(aPPX,PPX->PPX_ID)
           EndIf
       Endif    
       PPX->(dbSkip())
Enddo
ProcRegua(nCount)
For nZZ := 1 to nCount
    dbSelectArea("PPX")
    PPX->(dbSetOrder(1))
    PPX->(dbSeek(xFilial("PPX")+aPPX[nZZ]))
    If PPX->PPX_AGLUT $ "G|S|"
       Processa({|| PPX_PVPPM()},"Gerando Pedidos de Venda da PPM")
       Loop
    Endif   
    For nXX := 1 to 10000
        cNumPed := GetSxeNum('SC5','C5_NUM') 
        ConfirmSX8()
        dbSelectArea("SC5")
        SC5->(dbSetOrder(1))
        If !(SC5->(dbSeek(xFilial("SC5")+cNumPed)))
             nXX := 10000
        Endif
    Next nXX  
    cTexto  := ""      
    nItem   := 0
    aItens  := {}
    aCabec  := {}

   aCNPJ := fLerSA1(PPX->PPX_CNPJ)
   If Len(aCNPJ) = 0
      MsgStop("Cliente não cadastrado: "+PPX->PPX_CNPJ,"Erro")
      Return  
   Endif
   nRecno := aCNPJ[2]
   If nRecno = 0
      MsgStop("Cliente não cadastrado: "+PPX->PPX_CNPJ,"Erro")
   Return  
   Endif
      SA1->(dbGoTo(nRecno))  
      If SA1->A1_MSBLQL = "1"  
      MsgStop("Cliente Inativo: "+PPX->PPX_CNPJ,"Erro")
      Return
   Endif      
    
   IncProc(TIME()+"-Gravando PV "+cNumPed+" - NOR - "+SA1->A1_COD+"-"+SA1->A1_LOJA+" - "+Alltrim(SA1->A1_NREDUZ))
   aAdd(aCabec,{"C5_FILIAL" ,xFilial("SC5") 	,Nil})
   aAdd(aCabec,{"C5_NUM"	 ,cNumPed 	        ,Nil})
   aAdd(aCabec,{"C5_TIPO"   ,"N" 			    ,Nil})
   aAdd(aCabec,{"C5_CLIENTE",SA1->A1_COD		,Nil})
   aAdd(aCabec,{"C5_LOJACLI",SA1->A1_LOJA      ,Nil})
   aAdd(aCabec,{"C5_CLIENT" ,SA1->A1_COD       ,Nil})
   aAdd(aCabec,{"C5_LOJAENT",SA1->A1_LOJA      ,Nil})
   aAdd(aCabec,{"C5_TIPOCLI",SA1->A1_TIPO      ,Nil})
   aAdd(aCabec,{"C5_EMISSAO",PPX->PPX_EMISSA   ,Nil})
   aAdd(aCabec,{"C5_RPS"    ,PPX->PPX_RPS      ,Nil})
   aAdd(aCabec,{"C5_NFSE"	 ,PPX->PPX_NFSE     ,Nil})
   aAdd(aCabec,{"C5_TIPLIB" ,"1"               ,Nil})
   aAdd(aCabec,{"C5_TPCARGA","2"			    ,Nil})
   aAdd(aCabec,{"C5_OBS"	 ,PPX->PPX_ID+'-'+ALLTRIM(PPX->PPX_FILE),Nil})
   aAdd(aCabec,{"C5_NOMERED",SA1->A1_NREDUZ	,Nil})
   aAdd(aCabec,{"C5_CONDPAG",PPX->PPX_CONDPG   ,Nil})
   aAdd(aCabec,{"C5_NATUREZ",PPX->PPX_CODNAT   ,Nil})
   aAdd(aCabec,{"C5_XTPTAXA",PPX->PPX_AGLUT    ,Nil})
   aAdd(aCabec,{"C5_XDTVCTO",PPX->PPX_VENCTO   ,Nil})
   aAdd(aCabec,{"C5_XREDE"  ,fLerSX5("Z1",PPX->PPX_REDE)     ,Nil})
   aAdd(aCabec,{"C5_XMASTER",fLerSX5("Z2",PPX->PPX_MASTER)   ,Nil})
   dbSelectArea("PPZ")
   PPZ->(dbSetOrder(1))
   PPZ->(dbSeek(xFilial("PPZ")+PPX->PPX_ID))
   While !Eof() .And. PPZ->PPZ_FILIAL == xFilial('PPZ') .AND. PPZ->PPZ_ID = PPX->PPX_ID
         dbSelectArea("SB1")
         SB1->(dbSetOrder(1))
	     SB1->(dbSeek(xFilial("SB1")+PPZ->PPZ_CODPRO))
         axLinha := {}
         nItem ++
         aAdd(axLinha,{"C6_FILIAL" ,xFilial("SC6")                       ,Nil})
         aAdd(axLinha,{"C6_NUM"	,cNumPed                              ,Nil})
         aAdd(axLinha,{"C6_ITEM"   ,STRZERO(nItem,TamSx3("C6_ITEM")[1])  ,Nil})
         aAdd(axLinha,{"C6_PRODUTO",PPZ->PPZ_CODPRO                      ,Nil})
         aAdd(axLinha,{"C6_QTDVEN" ,PPZ->PPZ_QTDVEN                      ,Nil})
         aAdd(axLinha,{"C6_QTDLIB" ,PPZ->PPZ_QTDVEN                      ,Nil})
         aAdd(axLinha,{"C6_UM"     ,SB1->B1_UM                           ,Nil})      
         aAdd(axLinha,{"C6_PRCVEN" ,PPZ->PPZ_PRCVEN                      ,Nil})
         aAdd(axLinha,{"C6_PRUNIT" ,PPZ->PPZ_PRUNIT                      ,Nil})
         aAdd(axLinha,{"C6_VALOR"  ,PPZ->PPZ_QTDVEN*PPZ->PPZ_VALOR       ,Nil})
         aAdd(axLinha,{"C6_TPOP"   ,"F"                                  ,Nil})
         aAdd(axLinha,{"C6_SUGENTR",dDataBase                            ,Nil})
         aAdd(axLinha,{"C6_VALDESC",PPZ->PPZ_VLDESC                      ,Nil})
         aAdd(axLinha,{"C6_OPER"    ,"07"                                ,Nil})
         aAdd(axLinha,{"C6_TES"    ,PPZ->PPZ_TES                         ,Nil})
         aAdd(axLinha,{"C6_LOCAL"  ,PPZ->PPZ_LOCAL                       ,Nil})
         aAdd(axLinha,{"C6_CLI"    ,PPZ->PPZ_CLI                         ,Nil})
         aAdd(axLinha,{"C6_LOJA"   ,PPZ->PPZ_LOJA                        ,Nil})
         aAdd( aItens,axLinha )
         RecLock("PPZ",.F.)
         PPZ->PPZ_PEDIDO := cNumPed
         PPZ->(MsUnlock())
         PPZ->(dbSkip())
    Enddo
    //dbSelectArea("SC5")
    //RecLock("SC5",.T.)
    //For nX1 := 1 to LEN(aCabec)
	//    SC5->&(aCabec[nX1,1]) := aCabec[nX1,2]
    //Next nX1
    //SC5->(msUnlock())
    //For nX1 := 1 to LEN(aItens)
    //    dbSelectArea("SC6")
    //    RecLock("SC6",.T.)
    //    For nX2 := 1 to Len(aItens[nX1])
	//       SC6->&(aItens[nX1,nX2,1]) := aItens[nX1,nX2,2]
	//    Next nX2  
	//    SC6->(msUnlock()) 
    //Next nX1   
    lMsErroAuto := .F.
    MsExecAuto({|x,y,z| MATA410(x,y,z)},aCabec,aItens,3)
	If lMsErroAuto
	   If File( cStartPath+cArqLog )
          Ferase( cStartPath+cArqLog )
	   Endif
	   MostraErro(cStartPath,cArqLog)
	   If File(cStartPath+cArqLog )
          cLog := MemoRead(cStartPath+cArqLog)
          Disarmtransaction()
          Reclock("PPX",.F.)
	      PPX->PPX_STATUS := "0"
	      PPX->PPX_ERRO   := cLog
	      PPX->PPX_PEDIDO := " "
	      PPX->(MsUnlock())
	   Endif
	Else
       RecLock("PPX",.F.)
	   PPX->PPX_STATUS := "2"
	   PPX->PPX_PEDIDO := cNumPed
	   PPX->PPX_OK     := '  '
       PPX->(MsUnlock())
    Endif  
Next nZZ
cLog := "UPDATE "+RetSqlName("PPX")+" SET PPX_OK = '  '"
TCSQLExec(cLog)
Return

//-------------------------------------------------------------------------
Static Function PPX_PVPPM()     
//-------------------------------------------------------------------------
Local nX1        := 0
Local nX2        := 0
Local nItem      := 0
Local cNumPed    := ""
Local axLinha    := {}
Local aItens     := {}
Local aCabec     := {}
Local aPPX       := {}
Local nXX        := 0
Local nYY        := 0
Local nZZ        := 0
Local cTexto     := 0
Local cMarca     := oMark:Mark()
Local lInverte   := oMark:IsInvert()
Local nCount     := 0
Local nValor     := 0
Local nVezes     := 0
Local aTaxas     := {}
Local cParRede   := {}
Local aCNPJ      := {}
Local nRecno     := 0

cLog		 := ''
cArqLog    := 'LOGCSV.LOG'
aAdd(aTaxas,{PPX->PPX_TXADM,"ADM","Taxa Administrativa",""})
aAdd(aTaxas,{PPX->PPX_TXSRV,"SRV","Taxa de Serviços",""})
aAdd(aTaxas,{PPX->PPX_TXMID,"MID","Taxa de Mídia",""})

ProcRegua(2000)
Begin transaction
For nVezes := 1 to 3
cLog := ""
    For nXX := 1 to 10000
        cNumPed := GetSxeNum('SC5','C5_NUM') 
        ConfirmSX8()
        dbSelectArea("SC5")
        SC5->(dbSetOrder(1))
        If !(SC5->(dbSeek(xFilial("SC5")+cNumPed)))
           nXX := 10000
        Endif
    Next nXX  
   aCNPJ := fLerSA1(PPX->PPX_CNPJ)
   If Len(aCNPJ) = 0
      MsgStop("Cliente não cadastrado: "+PPX->PPX_CNPJ,"Erro")
      Return  
   Endif
   nRecno := aCNPJ[2]
   If nRecno = 0
      MsgStop("Cliente não cadastrado: "+PPX->PPX_CNPJ,"Erro")
   Return  
   Endif
      SA1->(dbGoTo(nRecno))  
      If SA1->A1_MSBLQL = "1"  
      MsgStop("Cliente Inativo: "+PPX->PPX_CNPJ,"Erro")
      Return
   Endif     
   cTexto   := fLerSX5("Z1",PPX->PPX_REDE)
   cParRede := LEFT(Alltrim(cTexto)+SPACE(30),TamSx3("PAR_E_REDE")[1])
   //dbSelectArea("SX5")
   //SX5->(dbSetOrder(1))
   //SX5->(dbSeek(xFilial("SX5")+cTexto))
   dbSelectArea("PAR")
   PAR->(dbSetOrder(2))
   If !(PAR->(dbSeek(xFilial("PAR")+aTaxas[nVezes,2]+cParRede+LEFT(Alltrim(PPX->PPX_PRODUT)+SPACE(30),TamSx3("PAR_PROD")[1]))))
      MsgStop("Não existe cadastro DE/PARA de Produtos por Rede-"+aTaxas[nVezes,2]+"-"+Alltrim(cParRede)+"-"+Alltrim(PPX->PPX_PRODUT),"Erro")
      Return
   Endif   
   DbSelectArea("SB1")
   SB1->(dbSetOrder(1))
   If !(SB1->(dbSeek(xFilial("SB1")+PAR->PAR_PRODTX)))
         MsgStop("Produto Procoop não cadastrado: "+aTaxas[nVezes,2]+"-"+Alltrim(cTexto)+"-"+Alltrim(PPX->PPX_PRODUT),"Erro")
      Return
   Endif  
   IncProc(TIME()+"-Gravando PV "+cNumPed+" - "+aTaxas[nVezes,2]+" - "+SA1->A1_COD+"-"+SA1->A1_LOJA+" - "+Alltrim(SA1->A1_NREDUZ))
   aCabec := {} 
   aAdd(aCabec,{"C5_FILIAL" ,xFilial("SC5") 	  ,Nil})
   aAdd(aCabec,{"C5_NUM"	,cNumPed 	          ,Nil})
   aAdd(aCabec,{"C5_TIPO"   ,"N" 			      ,Nil})
   aAdd(aCabec,{"C5_CLIENTE",SA1->A1_COD		  ,Nil})
   aAdd(aCabec,{"C5_LOJACLI",SA1->A1_LOJA         ,Nil})
   aAdd(aCabec,{"C5_CLIENT" ,SA1->A1_COD          ,Nil})
   aAdd(aCabec,{"C5_LOJAENT",SA1->A1_LOJA         ,Nil})
   aAdd(aCabec,{"C5_TIPOCLI",SA1->A1_TIPO         ,Nil})
   aAdd(aCabec,{"C5_EMISSAO",PPX->PPX_EMISSA      ,Nil})
   aAdd(aCabec,{"C5_RPS"    ,PPX->PPX_RPS         ,Nil})
   aAdd(aCabec,{"C5_NFSE"	,PPX->PPX_NFSE        ,Nil})
   aAdd(aCabec,{"C5_TIPLIB" ,"1"                  ,Nil})
   aAdd(aCabec,{"C5_TPCARGA","2"			      ,Nil})
   aAdd(aCabec,{"C5_OBS"	,PPX->PPX_ID+'-'+ALLTRIM(PPX->PPX_FILE),Nil})
   aAdd(aCabec,{"C5_NOMERED",SA1->A1_NREDUZ	      ,Nil})
   aAdd(aCabec,{"C5_CONDPAG",SB1->B1_XCOND        ,Nil})
   aAdd(aCabec,{"C5_NATUREZ",PAR->PAR_NATURZ      ,Nil})
   aAdd(aCabec,{"C5_XDTVCTO",PPX->PPX_VENCTO      ,Nil})
   aAdd(aCabec,{"C5_XTPTAXA",Left(aTaxas[nVezes,2],1),Nil})
   aAdd(aCabec,{"C5_XREDE"  ,PAR->PAR_E_REDE         ,Nil})
   aAdd(aCabec,{"C5_XMASTER",fLerSX5("Z2",PPX->PPX_MASTER) ,Nil})
   axLinha := {}
   aItens  := {}
   aAdd(axLinha,{"C6_FILIAL" ,xFilial("SC6")                 ,Nil})
   aAdd(axLinha,{"C6_NUM"	 ,cNumPed                        ,Nil})
   aAdd(axLinha,{"C6_ITEM"   ,STRZERO(1,TamSx3("C6_ITEM")[1]),Nil})
   aAdd(axLinha,{"C6_PRODUTO",SB1->B1_COD                    ,Nil})
   aAdd(axLinha,{"C6_QTDVEN" ,1                              ,Nil})
   aAdd(axLinha,{"C6_QTDLIB" ,1                              ,Nil})
   aAdd(axLinha,{"C6_UM"     ,SB1->B1_UM                     ,Nil})      
   aAdd(axLinha,{"C6_PRCVEN" ,aTaxas[nVezes,1]               ,Nil})
   aAdd(axLinha,{"C6_PRUNIT" ,aTaxas[nVezes,1]               ,Nil})
   aAdd(axLinha,{"C6_VALOR"  ,aTaxas[nVezes,1]               ,Nil})
   aAdd(axLinha,{"C6_TPOP"   ,"F"                            ,Nil})
   aAdd(axLinha,{"C6_SUGENTR",dDataBase                      ,Nil})
   aAdd(axLinha,{"C6_VALDESC",0                              ,Nil})
   aAdd(axLinha,{"C6_TES"    ,SB1->B1_TS                     ,Nil})
   aAdd(axLinha,{"C6_LOCAL"  ,SB1->B1_LOCPAD                 ,Nil})
   aAdd(axLinha,{"C6_CLI"    ,SA1->A1_COD                    ,Nil})
   aAdd(axLinha,{"C6_LOJA"   ,SA1->A1_LOJA                   ,Nil})   
   aAdd( aItens,axLinha )
   If aTaxas[nVezes,1] > 0
      cLog := ""
      //dbSelectArea("SC5")
      //RecLock("SC5",.T.)
      //For nX1 := 1 to LEN(aCabec)
	  //    SC5->&(aCabec[nX1,1]) := aCabec[nX1,2]
      //Next nX1
      //SC5->(msUnlock())
      //For nX1 := 1 to LEN(aItens)
      //    dbSelectArea("SC6")
      //    RecLock("SC6",.T.)
      //    For nX2 := 1 to Len(aItens[nX1])
	  //       SC6->&(aItens[nX1,nX2,1]) := aItens[nX1,nX2,2]
	  //    Next nX2  
	  //    SC6->(msUnlock()) 
      //Next nX1
      aTaxas[nVezes,4] := cNumPed
      If (PPX_GRVPV(cNumPed,3,Left(aTaxas[nVezes,3]+space(20),20),aCabec,aItens))
          cTexto := "Problemas na gravação do Pedido de Venda "+Chr(13)+Chr(10)
          cTexto += +Chr(13)+Chr(10)
          cTexto += cLog
         MsgStop (cTexto,"Erro")
         RecLock("PPX",.F.)
         PPX->PPX_ERRO := cTexto
         PPX->(msUnlock())
         Return
      Endif   
         IncProc(TIME()+"-Gerado PV "+cNumPed+" - "+aTaxas[nVezes,2]+" - "+SA1->A1_COD+"-"+SA1->A1_LOJA+" - "+Alltrim(SA1->A1_NREDUZ))
   Endif  
   dbSelectArea("SC5") 
   SC5->(dbSetOrder(1))
   SC5->(dbSeek(xFilial("SC5")+cNumPed))
   aTaxas[nVezes,4] := cNumPed
Next nVezes
cTexto := "Pedidos de Venda gerados"+Chr(13)+Chr(10)
For nVezes := 1 to 3
    cTexto += aTaxas[nVezes,4]+" ==> "+Left(aTaxas[nVezes,3]+Space(20),20)+Chr(13)+Chr(10)
Next nVezes
RecLock("PPX",.F.)
PPX->PPX_STATUS   := "2"
PPX->PPX_PEDIDO := aTaxas[1,4]
PPX->PPX_ERRO   := cTexto
PPX->(msUnlock())
End transaction
Return

//-------------------------------------------------------------------------
Static Function PPX_GRVPV(cNumero,nTipo,cTxDesc,aCabec,aItens)
//-------------------------------------------------------------------------
cLog		 := ''
cArqLog    := 'LOGCSV.LOG'

lMsErroAuto := .F.
MsExecAuto({|x,y,z| MATA410(x,y,z)},aCabec,aItens,nTipo)
If lMsErroAuto
   If File( cStartPath+cArqLog )
      Ferase( cStartPath+cArqLog )
	Endif
	Mostraerro(cStartPath,cArqLog)
	If File(cStartPath+cArqLog )
       cLog := MemoRead(cStartPath+cArqLog)
       Disarmtransaction()
       Reclock("PPX",.F.)
	   PPX->PPX_STATUS := "0"
	   PPX->PPX_ERRO   := cLog
	   PPX->PPX_PEDIDO := " "
	   PPX->(MsUnlock())
	Endif
Else
    dbSelectArea("PPZ")
    PPZ->(dbSetOrder(1))
    PPZ->(dbSeek(xFilial("PPZ")+PPX->PPX_ID))
    While !Eof() .And. PPZ->PPZ_FILIAL == xFilial('PPZ') .AND. PPZ->PPZ_ID = PPX->PPX_ID
          RecLock("PPZ",.F.)
          PPZ->PPZ_PEDIDO := cNumero
          PPZ->PPZ_EMPRES := PPX->PPX_EMPRES
          PPZ->(MsUnlock()) 
          PPZ->(dbSkip())
    Enddo      
Endif   
Return (lMsErroAuto)
  
//-------------------------------------------------------------------------
Static Function PPX_TELAEXCL(cPedNum,cID,axPedidos,pTipo)     
//-------------------------------------------------------------------------                   
Local oAtePedido
Local cAtePedido := PPX->PPX_PEDIDO
Local oButton1
Local oButton2
Local oDoPedido
Local cDoPedido := PPX->PPX_PEDIDO
Local oGroup1
Local oSay1
Local oSay2
Static oDlgExcl

cMsg := Space(2000)

DEFINE MSDIALOG oDlgExcl TITLE "Informe o intervalo de Números dos Pedidos de Venda" FROM 000, 000  TO 240, 390 COLORS 0, 16777215 PIXEL
    @ 001, 004 GROUP oGroup1 TO 028, 189 PROMPT "Informe os Números dos Pedidos de Venda" OF oDlgExcl COLOR 16711680, 16777215 PIXEL
    @ 015, 009 SAY oSay1 PROMPT "De :" SIZE 020, 007 OF oDlgExcl COLORS 0, 16777215 PIXEL
    @ 012, 032 MSGET oDoPedido VAR cDoPedido SIZE 060, 010 OF oDlgExcl COLORS 0, 16777215 PIXEL
    @ 015, 101 SAY oSay2 PROMPT "Até :" SIZE 020, 007 OF oDlgExcl COLORS 0, 16777215 PIXEL
    @ 012, 125 MSGET oAtePedido VAR cAtePedido SIZE 060, 010 OF oDlgExcl COLORS 0, 16777215 PIXEL
    @ 033, 004 GET oMsg VAR cMsg OF oDlgExcl MULTILINE SIZE 185, 064 COLORS 0, 16777215 HSCROLL NOBORDER PIXEL
     @ 102, 110 BUTTON oButton1 PROMPT "Confirmar" SIZE 037, 012 OF oDlgExcl ACTION PPX_PVEXCL(cDoPedido,cAtePedido,cID,axPedidos,pTipo) PIXEL 
    @ 102, 151 BUTTON oButton2 PROMPT "Sair" SIZE 037, 012 OF oDlgExcl ACTION oDlgExcl:End() PIXEL
ACTIVATE MSDIALOG oDlgExcl CENTERED
Return

//-------------------------------------------------------------------------
Static Function PPX_PVEXCL(cDoPedido,cAtePedido,cID,axPedidos,pTipo) 
//-------------------------------------------------------------------------

Local nXX     := 0
Local nCount  := 0
Local cQryAux := ""
Local aCabec  := {}
Local aItens  := {}
Local axLinha  := {}
Local aPPX    := {}
Local nTotal  := 0
Local nAtual  := 0

cMsg += "=> Processamento iniciado "+ Chr(13)+Chr(10)
cMsg += ""+ Chr(13)+Chr(10)
cPrefxTel = U_ETX_PREFIXO()
If Select("QRY_AUX")>0
   DbSelectArea("QRY_AUX")
   DbCloseArea()
Endif
oMsg:Refresh()
oMsg:GoEnd() 
oDlgExcl:Refresh()
cQryAux   := ""
cQryAux += " SELECT * FROM "+RetSqlName("PPX")+" PPX"
cQryAux += " INNER JOIN "+RetSqlName("SC5")+" SC5 ON C5_FILIAL = '"+xFilial("SC5")+"' AND LEFT(C5_OBS,"+Alltrim(Str(TamSx3('PPX_ID')[1]))+") = PPX_ID AND SC5.D_E_L_E_T_ = ' '"	"
cQryAux += " WHERE PPX.D_E_L_E_T_ = ' '"	
cQryAux += " AND PPX_FILIAL = '"+xFilial("PPX")+"'"	
cQryAux += " AND PPX_STATUS = '2'"
cQryAux += " AND PPX_PREFIX = '"+Alltrim(cPrefxTel)+"'"
If pTipo = 2
   cQryAux += " AND PPX_ID = '"+cID+"'"
Else   
   cQryAux += " AND PPX_PEDIDO BETWEEN '"+cDoPedido+"' AND '"+cAtePedido+"'"
Endif
cQryAux += " ORDER BY PPX_PEDIDO"	
cQryAux := ChangeQuery(cQryAux)
TCQuery cQryAux New Alias "QRY_AUX"                
Count to nCount 
dbselectarea("QRY_AUX")  
QRY_AUX->(DbGoTop())
ProcRegua(nCount)
While QRY_AUX->(!EOF())   
      nAtual++
      dbSelectArea("PPX")
      PPX->(dbSetOrder(1))
      PPX->(dbSeek(xFilial("PPX")+QRY_AUX->PPX_ID))
      dbSelectArea("SC5")
      SC5->(dbSetOrder(1))
      If (SC5->(dbSeek(xFilial("SC5")+QRY_AUX->C5_NUM)))
           aItens  := {}
           aCabec  := {}  
	       aAdd(aCabec,{"C5_FILIAL" ,xFilial("SC5"),Nil})
	       aAdd(aCabec,{"C5_NUM"	,SC5->C5_NUM    ,Nil})
	       aAdd(aCabec,{"C5_TIPO"   ,SC5->C5_TIPO   ,Nil})
	       aAdd(aCabec,{"C5_CLIENTE",SC5->C5_CLIENTE,Nil})
	       aAdd(aCabec,{"C5_LOJACLI",SC5->C5_LOJACLI,Nil})
	       aAdd(aCabec,{"C5_CLIENT" ,SC5->C5_CLIENT ,Nil})
	       aAdd(aCabec,{"C5_LOJAENT",SC5->C5_LOJAENT,Nil})
	       aAdd(aCabec,{"C5_TIPOCLI",SC5->C5_TIPOCLI,Nil})
	       aAdd(aCabec,{"C5_EMISSAO",SC5->C5_EMISSAO,Nil})
	       aAdd(aCabec,{"C5_TIPLIB" ,SC5->C5_TIPLIB ,Nil})
	       aAdd(aCabec,{"C5_TPCARGA",SC5->C5_TPCARGA,Nil})
           dbSelectArea("SC6")
           SC6->(dbSetOrder(1))
           SC6->(dbSeek(xFilial("SC6")+SC5->C5_NUM))
           While !Eof() .And. SC6->C6_FILIAL == xFilial('SC6') .AND. SC6->C6_NUM = SC5->C5_NUM
                 axLinha := {}
                 aAdd(axLinha,{"C6_FILIAL" ,xFilial("SC6") ,Nil})
                 aAdd(axLinha,{"C6_NUM"	,SC6->C6_NUM       ,Nil})
                 aAdd(axLinha,{"C6_ITEM"   ,SC6->C6_ITEM   ,Nil})
                 aAdd(axLinha,{"C6_PRODUTO",SC6->C6_PRODUTO,Nil})
                 aAdd(axLinha,{"C6_QTDVEN" ,SC6->C6_QTDVEN ,Nil})
                 aAdd(axLinha,{"C6_QTDLIB" ,0              ,Nil})
                 aAdd(axLinha,{"C6_UM"     ,SC6->C6_UM     ,Nil})      
                 aAdd(axLinha,{"C6_PRCVEN" ,SC6->C6_PRCVEN ,Nil})
                 aAdd(axLinha,{"C6_PRUNIT" , SC6->C6_PRCVEN,Nil})
                 aAdd(axLinha,{"C6_VALOR"  ,SC6->C6_PRCVEN  ,Nil})
                 aAdd(axLinha,{"C6_TPOP"   ,SC6->C6_TPOP   ,Nil})
                 aAdd(axLinha,{"C6_SUGENTR",SC6->C6_SUGENTR,Nil})
                 aAdd(axLinha,{"C6_VALDESC",0              ,Nil})
                 aAdd(axLinha,{"C6_TES"    ,SC6->C6_TES    ,Nil})
                 aAdd(axLinha,{"C6_LOCAL"  ,SC6->C6_LOCAL  ,Nil})
                 aAdd(axLinha,{"C6_CLI"    ,SC6->C6_CLI    ,Nil})
                 aAdd(axLinha,{"C6_LOJA"   ,SC6->C6_LOJA   ,Nil})
                 aAdd( aItens,axLinha )
                 SC6->(dbSkip())
           Enddo
           lMsErroAuto := .F.
           MsExecAuto({|x,y,z| MATA410(x,y,z)},aCabec,aItens,4)
		   If lMsErroAuto
		      If File( cStartPath+ cArqLog )
			     Ferase( cStartPath+ cArqLog )
		      Endif
		      Mostraerro(cStartPath, cArqLog)
		      cMsg += ">>>>> Problemas na Alteração do Pedido : "+PPX->PPX_ID+" - "+SC5->C5_NUM+ Chr(13)+Chr(10)
		      If File( cStartPath+cArqLog )
			     cLog := MemoRead(cStartPath+cArqLog)
			     cMsg += cLog+ Chr(13)+Chr(10)
		      Endif
		      oMsg:Refresh()
		      oMsg:GoEnd() 
              oDlgExcl:Refresh()	
              Disarmtransaction()
           Else   
              lMsErroAuto := .F.
              MsExecAuto({|x,y,z| MATA410(x,y,z)},aCabec,aItens,5)
		      If lMsErroAuto
		      If File( cStartPath+ cArqLog )
			     Ferase( cStartPath+ cArqLog )
		      Endif
		      Mostraerro(cStartPath, cArqLog)
		      cMsg += ">>>>> Problemas na Exclusão do Pedido : "+PPX->PPX_ID+" - "+SC5->C5_NUM+ Chr(13)+Chr(10)
		      If File( cStartPath+cArqLog )
			     cLog := MemoRead(cStartPath+cArqLog)
			     cMsg += cLog+ Chr(13)+Chr(10)
		      Endif
		         oMsg:Refresh()
		         oMsg:GoEnd() 
              oDlgExcl:Refresh()
                 Disarmtransaction()
		      Else
		         cMsg += "----> PV Excluído com sucesso ID Importação: "+PPX->PPX_ID+" - PV Número: "+SC5->C5_NUM+ Chr(13)+Chr(10)
		         oMsg:Refresh()
		         oMsg:GoEnd() 
              oDlgExcl:Refresh()
		         RecLock("PPX",.F.)
	             PPX->PPX_STATUS := "1"
	             PPX->PPX_OK     := '  '
	             PPX->PPX_ERRO   := ' '
                 PPX->(MsUnlock())
              Endif 
           Endif   
     Endif
     QRY_AUX->(DbSkip())      
EndDo
cMsg += ""+ Chr(13)+Chr(10)
cMsg += "=> Processamento finalizado "+ Chr(13)+Chr(10)
oMsg:Refresh()
oDlgExcl:Refresh()
Return

//-------------------------------------------------------------------------
Static Function PPX_TELALOTE(cDoPedido,cAtePedido)     
//-------------------------------------------------------------------------                   
Local oAtePedido
Local cAtePedido := PPX->PPX_ID
Local oButton1
Local oButton2
Local oDoPedido
Local cDoPedido  := PPX->PPX_ID
Local oSay1
Local oSay2
Static oDlgExcl

cMsg := Space(2000)

DEFINE MSDIALOG oDlgExcl TITLE "Informe o intervalo dos Id´s" FROM 000, 000  TO 240, 390 COLORS 0, 16777215 PIXEL
    @ 001, 004 GROUP oGroup1 TO 028, 189 PROMPT "Informe os Números dos ID´s" OF oDlgExcl COLOR 16711680, 16777215 PIXEL
    @ 015, 009 SAY oSay1 PROMPT "De :" SIZE 020, 007 OF oDlgExcl COLORS 0, 16777215 PIXEL
    @ 012, 032 MSGET oDoPedido VAR cDoPedido SIZE 060, 010 OF oDlgExcl COLORS 0, 16777215 PIXEL
    @ 015, 101 SAY oSay2 PROMPT "Até :" SIZE 020, 007 OF oDlgExcl COLORS 0, 16777215 PIXEL
    @ 012, 125 MSGET oAtePedido VAR cAtePedido SIZE 060, 010 OF oDlgExcl COLORS 0, 16777215 PIXEL
    @ 033, 004 GET oMsg VAR cMsg OF oDlgExcl MULTILINE SIZE 185, 064 COLORS 0, 16777215 HSCROLL NOBORDER PIXEL 
    @ 102, 110 BUTTON oButton1 PROMPT "Confirmar" SIZE 037, 012 OF oDlgExcl ACTION PPX_LOTE(cDoPedido,cAtePedido) PIXEL 
    @ 102, 151 BUTTON oButton2 PROMPT "Sair" SIZE 037, 012 OF oDlgExcl ACTION oDlgExcl:End() PIXEL
ACTIVATE MSDIALOG oDlgExcl CENTERED
Return

//-------------------------------------------------------------------------
Static Function PPX_LOTE(cDoPedido,cAtePedido)
//-------------------------------------------------------------------------
Local nXX     := 0
Local nCount  := 0
Local cQryAux := ""
Local aCabec  := {}
Local aItens  := {}
Local axLinha  := {}
Local aPPX    := {}
Local nTotal  := 0
Local nAtual  := 0

cPrefxTel := U_ETX_PREFIXO()
cMsg += "=> Processamento iniciado "+ Chr(13)+Chr(10)
cMsg += ""+Chr(13)+Chr(10)
If Select("QRY_AUX")>0
   DbSelectArea("QRY_AUX")
   DbCloseArea()
Endif
oMsg:Refresh()
oMsg:GoEnd() 
oDlgExcl:Refresh()
cQryAux   := ""
cQryAux += " SELECT * FROM "+RetSqlName("PPX")
cQryAux += " WHERE D_E_L_E_T_ = ''"
cQryAux += " AND PPX_FILIAL = '"+xFilial("PPX")+"'"
cQryAux += " AND PPX_STATUS IN ('0','1')"
cQryAux += " AND PPX_ID BETWEEN '"+cDoPedido+"' AND '"+cAtePedido+"'"
cQryAux += " AND PPX_PEDIDO = ' '"
cQryAux += " AND PPX_PREFIX = '"+Alltrim(cPrefxTel)+"'"
cQryAux += " ORDER BY PPX_ID"
cQryAux := ChangeQuery(cQryAux)
TCQuery cQryAux New Alias "QRY_AUX"
Count to nTotal
ProcRegua(nTotal)
QRY_AUX->(DbGoTop())
DbSelectArea("QRY_AUX")
While !QRY_AUX->(Eof())
      nAtual++
      dbSelectArea("PPX")
      PPX->(dbSetOrder(1))
      PPX->(dbSeek(xFilial("PPX")+QRY_AUX->PPX_ID))
      cMsg += "----> Registro importado, excluído com sucesso : "+PPX->PPX_ID+" - "+PPX->PPX_FANTAS+ Chr(13)+Chr(10)
      dbSelectArea("PPZ")
      PPZ->(dbSetOrder(1))
      PPZ->(dbSeek(xFilial("PPZ")+PPX->PPX_ID))
      While !PPZ->(Eof()) .AND. (PPZ->PPZ_FILIAL+PPZ->PPZ_ID == PPX->PPX_FILIAL+PPX->PPX_ID)
            RecLock("PPZ",.F.)
            PPZ->(dbDelete())
            PPZ->(MsUnlock())
            PPZ->(dbSkip())
      Enddo 
      RecLock("PPX",.F.)
      PPX->(dbDelete())
      PPX->(MsUnlock())     
      oMsg:Refresh()
      oMsg:GoEnd() 
      oDlgExcl:Refresh()
      QRY_AUX->(DbSkip())      
EndDo
cMsg += ""+ Chr(13)+Chr(10)
cMsg += "=> Processamento finalizado "+ Chr(13)+Chr(10)
oMsg:Refresh()
oDlgExcl:Refresh()

cQryAux := "DELETE FROM "+RetSqlName("PPX")+" WHERE D_E_L_E_T_ <> ' '"
TCSQLExec(cQryAux)

cQryAux := "DELETE FROM "+RetSqlName("PPZ")+" WHERE D_E_L_E_T_ <> ' '"
TCSQLExec(cQryAux)
Return

///-------------------------------------------------------------------------
Static Function PPX_NFGERA(p_par,cId)
///-------------------------------------------------------------------------
Local oButton1
Local oButton2
Local oCboTipo
Local oDoPedido
Local cDoPedido  := PPX->PPX_PEDIDO
Local oAtePedido
Local cAtePedido := PPX->PPX_PEDIDO
Local oDtImp
Local dDtImp     := PPX->PPX_DATA
Local oPeriodo
Local cPeriodo   := PPX->PPX_PERIOD
Local oButton1
Local oButton2
Local oCboTipo
Local nCboTipo   := 4
Local aCboTipo   := {"1=Administração","2=Serviços","3=Mídia","4=Normal","5=Infração","Churrascarias"}
Local oGroup1
Local oSay1
Local oSay2
Local oSay3
Local oSay4
Local oSay5
Local oSerieNF
Local cSerieNF := Space(TamSx3('F2_SERIE')[1])
Local aAglut  :={"1=Administração","2=Serviços","3=Mídia","4=Normal","5=Infração"}
Local nMes    := Month(dDataBase)-1
Local nAno    := Year(dDataBase)-1
Local cTitulo := If(p_par=1,"Seleção de Pedidos de Venda","Seleção de NF de Saída")
Local cTitPed := If(p_par=1,"Pedido De/Até:","NF De/Até:")

If p_par = 2
   cDoPedido  := Space(TamSx3('F2_DOC')[1])
   cAtePedido := Space(TamSx3('F2_DOC')[1])
Endif
If nMes < Month(dDataBase)
   nAno := nAno + 1   
Endif   
cMsg := Space(2000)

DEFINE MSDIALOG oDlgNFC TITLE cTitulo FROM 000, 000  TO 220, 410 COLORS 0, 16777215 PIXEL
       @ 001, 002 GROUP oGroup1 TO 106, 202 PROMPT "Informe os dados abaixo" OF oDlgNFC COLOR 16711680, 16777215 PIXEL
       If p_par > 1
          @ 015, 008 SAY oSay1 PROMPT "Tipo de Operação:" SIZE 050, 007 OF oDlgNFC COLORS 16711680, 16777215 PIXEL
          @ 015, 059 MSCOMBOBOX oCboTipo VAR nCboTipo ITEMS aCboTipo SIZE 072, 010 OF oDlgNFC COLORS 0, 16777215 PIXEL
       Endif
       @ 030, 008 SAY oSay2 PROMPT "Período:" SIZE 025, 007 OF oDlgNFC COLORS 16711680, 16777215 PIXEL
       @ 030, 059 MSGET oPeriodo VAR cPeriodo SIZE 060, 010 OF oDlgNFC COLORS 0, 16777215 PIXEL
       @ 045, 008 SAY oSay3 PROMPT "Data Importação:" SIZE 050, 007 OF oDlgNFC COLORS 16711680, 16777215 PIXEL
       @ 045, 059 MSGET oDtImp VAR dDtImp SIZE 060, 010 OF oDlgNFC COLORS 0, 16777215 PIXEL
       If p_par > 1
          @ 045, 127 SAY oSay5 PROMPT "Série NF:" SIZE 025, 007 OF oDlgNFC COLORS 16711680, 16777215 PIXEL
          @ 045, 158 MSGET oSerieNF VAR cSerieNF SIZE 033, 010 OF oDlgNFC COLORS 0, 16777215 PIXEL
       Endif
       @ 060, 008 SAY oSay4 PROMPT cTitPed SIZE 050, 007 OF oDlgNFC COLORS 16711680, 16777215 PIXEL
       @ 060, 059 MSGET oDoPedido VAR cDoPedido SIZE 060, 010 OF oDlgNFC COLORS 0, 16777215 PIXEL
       @ 060, 131 MSGET oAtePedido VAR cAtePedido SIZE 060, 010 OF oDlgNFC COLORS 0, 16777215 PIXEL
       @ 088, 122 BUTTON oButton1 PROMPT "Selecionar" SIZE 037, 012 OF oDlgNFC ACTION PPX_NFCGERA(cPeriodo,nCboTipo,dDtImp,cDoPedido,cAtePedido,cSerieNF,p_par,cId) PIXEL
       @ 089, 162 BUTTON oButton2 PROMPT "Sair" SIZE 037, 012 OF oDlgNFC ACTION oDlgNFC:End() PIXEL
 ACTIVATE MSDIALOG oDlgNFC CENTERED
Return

//-------------------------------------------------------------------------=
Static Function PPX_NFCGERA(cPeriodo,nCboTipo,dDtImp,cDoPedido,cAtePedido,cSerieNF,p_par,cId)
//-------------------------------------------------------------------------=               
Local oAtePedido
Local cAtePedido := cAtePedido
Local oButton1
Local oButton2
Local oCboTipo
Local nCboTipo := nCboTipo
Local aCboTipo := {"1=Administração","2=Serviços","3=Mídia","4=Normal","5=Infração","Churrascarias"}
Local oDoPedido
Local cDoPedido := cDoPedido
Local oDtImp
Local dDtImp := dDtImp
Local oGroup1
Local oPeriodo
Local oSay1
Local oSay2
Local oSay3
Local oSay4
Local oSay5
Local oSerieNF
Local aAglut  := {"1=Administração","2=Serviços","3=Mídia","4=Normal","5=Infração"}
Local cTitulo := If(p_par=1,"Seleção de Pedidos de Venda","Seleção de NF de Saída")
Local cTitPed := If(p_par=1,"Pedido De/Até:","NF De/Até:")


cMsg := Space(2000)

DEFINE MSDIALOG oDlgNFS TITLE cTitulo FROM 000, 000  TO 500, 1000 COLORS 0, 16777215 PIXEL  
    @ 001, 002 GROUP oGroup1 TO 084, 202 PROMPT "Informe os dados abaixo" OF oDlgNFS COLOR 16711680, 16777215 PIXEL
    @ 030, 008 SAY oSay2 PROMPT "Período:" SIZE 025, 007 OF oDlgNFS COLORS 16711680, 16777215  PIXEL
    @ 030, 059 MSGET oPeriodo VAR cPeriodo SIZE 060, 010 OF oDlgNFS COLORS 0, 16777215 READONLY PIXEL
    @ 045, 008 SAY oSay3 PROMPT "Data Importação:" SIZE 050, 007 OF oDlgNFS COLORS 16711680, 16777215 PIXEL
    @ 045, 059 MSGET oDtImp VAR dDtImp SIZE 060, 010 OF oDlgNFS COLORS 0, 16777215 READONLY PIXEL
    @ 089, 002 GET oMsg VAR cMsg OF oDlgNFS MULTILINE SIZE 200, 128 COLORS 0, 16777215 HSCROLL PIXEL
    @ 060, 008 SAY oSay4 PROMPT cTitPed SIZE 050, 007 OF oDlgNFS COLORS 16711680, 16777215 PIXEL
    @ 060, 059 MSGET oDoPedido VAR cDoPedido SIZE 060, 010 OF oDlgNFS COLORS 0, 16777215 READONLY PIXEL
    @ 060, 131 MSGET oAtePedido VAR cAtePedido SIZE 060, 010 OF oDlgNFS COLORS 0, 16777215 READONLY PIXEL
    @ 001, 203 GROUP oGroup3 TO 217, 495 PROMPT "Marque/Desmarque os itens desejados" OF oDlgNFS COLOR 16711680, 16777215 PIXEL 
    @ 229, 413 BUTTON oButton1 PROMPT "Confirmar" SIZE 037, 012 OF oDlgNFS ACTION Processa({|| (PPX_NFGRAVA(cPeriodo,nCboTipo,dDtImp,cDoPedido,cAtePedido,cSerieNF,p_par,cId))},"Gerando Nota Fiscal") PIXEL
    @ 229, 456 BUTTON oButton2 PROMPT "Sair" SIZE 037, 012 OF oDlgNFS ACTION fSair() PIXEL
    fWBrowse1(cPeriodo,nCboTipo,dDtImp,cDoPedido,cAtePedido,cSerieNF,p_par) 
ACTIVATE MSDIALOG oDlgNFS CENTERED
Return

///-------------------------------------------------------------------------
Static Function fSair()
///-------------------------------------------------------------------------
oDlgNFS :End()
oDlgNFC :End()
Return

///-------------------------------------------------------------------------
Static Function fWBrowse1(cPeriodo,nCboTipo,dDtImp,cDoPedido,cAtePedido,cSerieNF,p_par)
///-------------------------------------------------------------------------
Local oOk       := LoadBitmap( GetResources(), "LBOK")
Local oNo       := LoadBitmap( GetResources(), "LBNO")
Local cQryAux   := ""
Local cTpTaxa   := ""
Local nTpTaxa   := 0
Local nTotal    := 0
Local aAglut    := {"A","S","M","N","I"}

cPrefxTel := U_ETX_PREFIXO()
aWBrowse1 := {}
nTpTaxa := If(ValType(nCboTipo)="C",Val(nCboTipo),nCboTipo)
cTpTaxa := aAglut[nTpTaxa]
If Select("QRY_AUX")>0
   DbSelectArea("QRY_AUX")
   DbCloseArea()
Endif
If p_par = 1
   cQryAux += " SELECT * FROM "+RetSqlName("SC5")+" SC5"
   cQryAux += " INNER JOIN "+RetSqlName("PPX")+" PPX ON PPX.D_E_L_E_T_ = ' ' AND PPX_STATUS IN ('2','4')"
   cQryAux += "                                AND LEFT(C5_OBS,6) = PPX_ID AND PPX_PERIOD = '"+cPeriodo+"' AND PPX_PREFIX = '"+Alltrim(cPrefxTel)+"'" 
   cQryAux += "                                AND PPX_DATA ='"+DTOS(dDtImp)+"' "
   cQryAux += "                                AND PPX_EMISSA = C5_EMISSAO C*O*L*L*A*T*E*"
   cQryAux += " WHERE SC5.D_E_L_E_T_ = ''"
   cQryAux += " AND C5_FILIAL = '"+xFilial("SC5")+"'"
   cQryAux += " AND C5_LIBEROK = 'S'" 
   cQryAux += " AND C5_NOTA = ' '" 
   cQryAux += " AND C5_NUM BETWEEN '"+cDoPedido+"' AND '"+cAtePedido+"'"
   cQryAux += " ORDER BY C5_NUM"
Endif
If p_par = 2
   cQryAux += " SELECT * FROM "+RetSqlName("SC5")+" SC5"
   cQryAux += " JOIN "+RetSqlName("PPX")+" PPX ON PPX.D_E_L_E_T_ = ' ' AND PPX_STATUS IN ('4','5')"
   cQryAux += "                                AND LEFT(C5_OBS,6) = PPX_ID AND PPX_PERIOD = '"+cPeriodo+"'" 
   cQryAux += "                                AND PPX_DATA ='"+DTOS(dDtImp)+"' "
   cQryAux += "                                AND PPX_EMISSA = C5_EMISSAO C*O*L*L*A*T*E*"
   cQryAux += " JOIN "+RetSqlName("SF2")+" SF2 ON SF2.D_E_L_E_T_ = ' ' AND F2_CLIENTE = C5_CLIENTE C*O*L*L*A*T*E* AND F2_LOJA = C5_LOJACLI C*O*L*L*A*T*E*"
   cQryAux += "                             AND F2_DOC = C5_NOTA C*O*L*L*A*T*E* AND F2_SERIE = C5_SERIE C*O*L*L*A*T*E*"
   cQryAux += " WHERE SC5.D_E_L_E_T_ = ''"
   cQryAux += " AND C5_FILIAL = '"+xFilial("SC5")+"'"
   cQryAux += " AND C5_NOTA BETWEEN '"+cDoPedido+"' AND '"+cAtePedido+"'"
   cQryAux += " AND C5_SERIE = '"+cSerieNF+"'"
   cQryAux += " AND C5_XTPTAXA = '"+cTpTaxa+"'"
   cQryAux += " ORDER BY C5_NOTA"
Endif
If p_par = 3
   cQryAux += " SELECT * FROM "+RetSqlName("SC5")+" SC5"
   cQryAux += " JOIN "+RetSqlName("PPX")+" PPX ON PPX.D_E_L_E_T_ = ' ' AND PPX_STATUS IN ('4','5')"
   cQryAux += "                                AND LEFT(C5_OBS,6) = PPX_ID AND PPX_PERIOD = '"+cPeriodo+"'" 
   cQryAux += "                                AND PPX_DATA ='"+DTOS(dDtImp)+"' "
   cQryAux += "                                AND PPX_EMISSA = C5_EMISSAO C*O*L*L*A*T*E*"
   cQryAux += " JOIN "+RetSqlName("SF2")+" SF2 ON SF2.D_E_L_E_T_ = ' ' AND F2_CLIENTE = C5_CLIENTE C*O*L*L*A*T*E* AND F2_LOJA = C5_LOJACLI C*O*L*L*A*T*E*"
   cQryAux += "                             AND F2_DOC = C5_NOTA C*O*L*L*A*T*E* AND F2_SERIE = C5_SERIE C*O*L*L*A*T*E*"
   cQryAux += " WHERE SC5.D_E_L_E_T_ = ''"
   cQryAux += " AND C5_FILIAL = '"+xFilial("SC5")+"'"
   cQryAux += " AND C5_NOTA BETWEEN '"+cDoPedido+"' AND '"+cAtePedido+"'"
   cQryAux += " AND C5_SERIE = '"+cSerieNF+"'"
   cQryAux += " AND C5_XTPTAXA = '"+cTpTaxa+"'"
   cQryAux += " ORDER BY C5_NOTA"
Endif
cQryAux := If(cEmpAnt="99",Replace(cQryAux,"C*O*L*L*A*T*E*","COLLATE Latin1_General_CI_AS"),Replace(cQryAux,"C*O*L*L*A*T*E*",""))
cQryAux := ChangeQuery(cQryAux)
TCQuery cQryAux New Alias "QRY_AUX"
QRY_AUX->(DbGoTop())
DbSelectArea("QRY_AUX")
While !QRY_AUX->(Eof())	
      aItens   := {}
      aPvlNfs  := {}
      dbSelectArea("PPX")
      PPX->(dbSetOrder(1))
      PPX->(dbSeek(xFilial("PPX")+QRY_AUX->PPX_ID))
      
      dbSelectArea("SC5")
      SC5->(dbSetOrder(1))
      SC5->(dbSeek(QRY_AUX->C5_FILIAL+QRY_AUX->C5_NUM))
   
      dbSelectArea("SA1")
      SA1->(dbSetOrder(1))
      SA1->(dbSeek(xFilial("SA1")+SC5->C5_CLIENTE+SC5->C5_LOJACLI))
    
      nTotal := 0
      cQryAux := "Normal"
      dbSelectArea("SC6")
      SC6->(dbSetOrder(1))
      SC6->(dbSeek(xFilial("SC6")+SC5->C5_NUM))
      If SC5->C5_XTPTAXA = "A"
         cQryAux := "Administração"
      Endif   
            If SC5->C5_XTPTAXA = "M"
         cQryAux := "Mídia"
      Endif 
      If SC5->C5_XTPTAXA = "S"
         cQryAux := "Serviço"
      Endif 
      If SC5->C5_XTPTAXA = "N"
         cQryAux := "Normal"
      Endif 
      If SC5->C5_XTPTAXA = "I"
         cQryAux := "Infração"
      Endif 
            If SC5->C5_XTPTAXA = "C"
         cQryAux := "Churrascarias"
      Endif 
      While !SC6->(Eof()) .AND. SC6->C6_NUM == SC5->C5_NUM	  
             nTotal += SC6->C6_VALOR
             SC6->(dbSkip())
      Enddo 
      If p_par = 1
         aAdd(aWBrowse1,{.T.,SC5->C5_NUM,PPX->PPX_ID,cQryAux,SC5->C5_NOTA,SC5->C5_CLIENTE,SA1->A1_NREDUZ,Padr(TRANSFORM(nTotal,"@E 9,999,999,999.99"),16)})
      Else   
         aAdd(aWBrowse1,{.T.,SC5->C5_NUM,PPX->PPX_ID,cQryAux,SF2->F2_DOC ,SC5->C5_CLIENTE,SA1->A1_NREDUZ,Padr(TRANSFORM(nTotal,"@E 9,999,999,999.99"),16)})
     Endif    
      QRY_AUX->(dbSkip())
 Enddo
 aAdd(aWBrowse1,{.F.," "," "," "," "," ","",0})
lFirst := .F.
@ 013, 207 LISTBOX oWBrowse1 Fields HEADER "","Pedido","ID","Taxa","Num NF","Cliente","Fantasia","Valor" SIZE 282, 197 OF oDlgNFS PIXEL ColSizes 50,50
    oWBrowse1:SetArray(aWBrowse1)
    oWBrowse1:bLine := {|| {;
      If(aWBrowse1[oWBrowse1:nAT,1],oOk,oNo),;
      aWBrowse1[oWBrowse1:nAt,2],;
      aWBrowse1[oWBrowse1:nAt,3],;
      aWBrowse1[oWBrowse1:nAt,4],;
      aWBrowse1[oWBrowse1:nAt,5],;
      aWBrowse1[oWBrowse1:nAt,6],;
      aWBrowse1[oWBrowse1:nAt,7],;
      aWBrowse1[oWBrowse1:nAt,8]}}
    oWBrowse1:bLDblClick := {|| aWBrowse1[oWBrowse1:nAt,1] := !aWBrowse1[oWBrowse1:nAt,1],;
    oWBrowse1:DrawSelect()}
Return

///-------------------------------------------------------------------------
Static Function PPX_NFGRAVA(cPeriodo,nCboTipo,dDtImp,cDoPedido,cAtePedido,cSerieNF,p_par,cId)
///-------------------------------------------------------------------------
Local nX1      := 0
Local nX2      := 0
Local cTexto   := ""
Local cNumNF   := ""
Local cQryAux  := ""
Local cTpTaxa  := ""
Local nTpTaxa  := 0
Local aAglut   := {"A"  ,"S"  ,"M"  ,"I"  ,"N","C"}
Local aTpSer   := {"ADM","SER","MID","INF","NOR","CHU"}
Local nLinhas   := 0 
Local aPvlNfs   := {}
Local aCabec    := {}
Local aItens    := {}
Local aNotas    := {}
Local cNumNF    := ""
Local nQtdeNf   := 0
Local aParam460 := Array(30)
Local cxTPTaxa  := "N"
Local dxUltData := CTOD("//")
Local lOK       := .F.
Local aSays      := {}
Local aBotoes    := {}
Local oWnd
Local nxpOpc     := 0
 
If p_par = 2
   Processa({|| PPX_EXCNF(cPeriodo,nCboTipo,dDtImp,cDoPedido,cAtePedido,cSerieNF,p_par)},"Excluindo Notas Fiscais")
   Return         
Endif   
If p_par = 3
   Processa({|| PPX_TRANSMNF(cPeriodo,nCboTipo,dDtImp,cDoPedido,cAtePedido,cSerieNF,p_par)},"Excluindo Notas Fiscais")
   Return         
Endif  
dxUltData := fUltimaData()
aSays := {}
dDataBase := dDBaseBKP
aAdd(aSays,"Data Base       : "+DTOC(dDataBase))
aAdd(aSays,"Data do arquivo : "+DTOC(PPX->PPX_EMISSA))
aAdd(aSays,"Data última NF  : "+DTOC(dxUltData)) 
aAdd(aSays,"Se optar por OK, toda a Data de Emissão do arquivo deverá ser "+DTOC(PPX->PPX_EMISSA))            
aAdd(aSays,"Se optar por PARAM, as NFs serão geradas com a Data Base "+DTOC(dDataBase))   
If DTOS(dxUltData)  > DTOS(PPX->PPX_EMISSA) 
   aAdd(aSays,"***** A T E N Ç Ã O *****")
   aAdd(aSays,"====> Última NF emitida foi em "+DTOC(dxUltData)+", portanto, as NFs serão geradas com a Data Base "+DTOC(dDataBase))
Endif
aAdd(aBotoes, {1, .T., {|o| nxpOpc := 1,o:oWnd:End()}})
aAdd(aBotoes, {5, .T., {|o| nxpOpc := 2,o:oWnd:End()}})
aAdd(aBotoes, {2, .T., {|o| nxpOpc := 3,o:oWnd:End()}})
FormBatch("Escolha a Data de Emissão das Notas Fiscais",aSays,aBotoes,,,650)

If nxpOpc = 3
   Return
Endif   
If nxpOpc = 1 .AND. DTOS(PPX->PPX_EMISSA) >= DTOS(dxUltData)
   dDataBase := PPX->PPX_EMISSA
Endif   
cEmpAnt   := cEmpBKP  
cFilAnt   := cFilBKP 
OpenFile(cEmpAnt+cFilAnt)
OpenSM0()
For nX1 := 1 to 30
		aParam460[nX1] := &("MV_PAR"+StrZero(nX1,2))
	Next nX1 
cMsg := Time()+"--> Iniciando a geração de Documentos de Saída (NF) "+Chr(13)+Chr(10)
oMsg:Refresh()
oMsg:GoEnd() 
oDlgNFS:Refresh()
nTpTaxa := If(ValType(nCboTipo)="C",Val(nCboTipo),nCboTipo)
cTpTaxa := aAglut[nTpTaxa]

ProcRegua(Len(aWBrowse1))
For nX1 := 1 to Len(aWBrowse1)
    If aWBrowse1[nX1,1]
       aItens   := {}
       aPvlNfs  := {}
       dbSelectArea("SC5")
       SC5->(dbSetOrder(1))
       SC5->(dbSeek(xFilial("SC5")+aWBrowse1[nX1,2]))
       dbSelectArea("PPX")
       PPX->(dbSetOrder(1))
       PPX->(dbSeek(xFilial("PPX")+LEFT(SC5->C5_OBS,6)))
       dbSelectArea("SE4")
       SE4->(dbSetOrder(1))
       SE4->(dbSeek(xFILIAL("SE4")+SC5->C5_CONDPAG))
       dbSelectArea("SC9")
       SC9->(dbSetOrder(1))
       SC9->(dbSeek(SC5->C5_FILIAL+SC5->C5_NUM))
       While !SC9->(Eof()) .AND. SC9->C9_FILIAL+SC9->C9_PEDIDO==SC5->C5_FILIAL+SC5->C5_NUM 
             If SC9->C9_FILIAL+SC9->C9_PEDIDO==SC5->C5_FILIAL+SC5->C5_NUM 	
                If Alltrim(SC9->C9_BLCRED)+Alltrim(SC9->C9_BLEST) = ""
                   dbSelectArea("SC6")
                   SC6->(dbSetOrder(1))
                   SC6->(dbSeek(SC9->C9_FILIAL+SC9->C9_PEDIDO+SC9->C9_ITEM+SC9->C9_PRODUTO))
                   dbSelectArea("SF4")
                   SF4->(dbSetOrder(1))
                   SF4->(dbSeek(xFILIAL("SF4")+SC6->C6_TES))
                   FG_Seek("SB1","SC9->C9_PRODUTO",1)
                   FG_Seek("SC5","SC9->C9_PEDIDO",1,.F.)
                   FG_Seek("SC6","SC9->C9_PEDIDO+SC9->C9_ITEM",1)
                   FG_Seek("SB5","SB1->B1_COD")
                   FG_Seek("SB2","SB1->B1_COD")
                   aItens := Array(14)
                   aItens[01] := SC9->C9_PEDIDO
                   aItens[02] := SC9->C9_ITEM
                   aItens[03] := SC9->C9_SEQUEN
                   aItens[04] := SC9->C9_QTDLIB
                   aItens[05] := SC9->C9_PRCVEN
                   aItens[06] := SC9->C9_PRODUTO                      
                   aItens[07] := SF4->F4_ISS=="N"
                   aItens[08] := SC9->(RecNo())
                   aItens[09] := SC5->(RecNo())
                   aItens[10] := SC6->(RecNo())
                   aItens[11] := SE4->(RecNo())
                   aItens[12] := SB1->(RecNo())
                   aItens[13] := SB2->(RecNo())
                   aItens[14] := SF4->(RecNo())
                   aAdd(aPvlNfs,aItens)     
                Endif           
             Endif
             SC9->(DbSkip())  
       Enddo      
       If Len(aPvlNfs) > 0  
       cxTPTaxa := SC5->C5_XTPTAXA
          nX3 := aScan(aAglut,Alltrim(cxTPTaxa))
          If nX3 = 0 .OR. nX3 > Len(aAglut) 
             cxTPTaxa := "N"
          Endif 
          nX3 := aScan(aAglut,Alltrim(cxTPTaxa))
          If nX3 = 0 .OR. nX3 > Len(aAglut) 
             nX3 := 4
          Endif 
          nX2 := aScan(aMVFATSER,Alltrim(aTpSer[nX3])) 
          If nX2 > 0 
             Begin Transaction
                   cSerieNF := Substr(aMVFATSER[nX2],5,TamSx3('F2_SERIE')[1])  
                   cNumNf := MaPvlNfs(aPvlNfs,cSerieNF,.F.,.F.,.F.,.T.,.F.,0,0,.T.,.F.) 
                   Sleep(500)
                   IncProc("NF: "+cNumNf+"-"+cSerieNF+" Pedido: "+SC5->C5_NUM+"-"+PPX->PPX_FANTAS)
                   If Alltrim(cNumNf) <> ""
                      nQtdeNf ++
                      aAdd(aNotas,{cNumNf,cSerieNF,SC5->C5_NUM,SC5->C5_CLIENTE,SC5->C5_LOJACLI,PPX->PPX_ID})  
                      cMsg += Time()+"--> NF gerada "+cNumNf+"-"+cSerieNF+" refer ao PV "+SC5->C5_NUM+Chr(13)+Chr(10)
                      dbSelectArea("SE1")
                      SE1->(dbSetOrder(1))
                      If SE1->(dbSeek(xFilial("SE1")+LEFT(aLLTRIM(cSerieNF)+SPACE(10),TamSx3('E1_PREFIXO')[1])+Left(cNumNf+Space(10),TamSx3('E1_NUM')[1])+Space(TamSx3('E1_PARCELA')[1])+LEFT("NF"+SPACE(10),TamSx3('E1_TIPO')[1])))
                         RecLock("SE1",.F.)
                         SE1->E1_EMISSAO := PPX->PPX_EMISSA
                         SE1->E1_VENCTO  := PPX->PPX_VENCTO
                         SE1->E1_VENCREA := PPX->PPX_VENCTO
                         SE1->(msUnlock())
                      Endif   
                      RecLock("PPX",.F.)
                      PPX->PPX_STATUS = "4"
                      PPX->PPX_ERRO += Chr(13)+Chr(10)+"NF "+cNumNf+"-"+cSerieNF+" gerada para o PV "+SC5->C5_NUM
                     PPX->(msUnlock())
                   Endif   
                   fLegFat(PPX->PPX_ID)
             End Transaction
             oMsg:Refresh()
             oMsg:GoEnd() 
             oDlgNFS:Refresh()
          Endif 
       Endif
    Endif 
Next nX1
For nX1 := 1 to Len(aNotas)
    fLegFat(aNotas[nX1,6])
Next nX1
cMsg += Time()+"--> "+Alltrim(Str(nQtdeNf))+" Documentos gerados"+Chr(13)+Chr(10)
cMsg += Time()+"--> Finalizado a geração de Documentos de Saída (NF)"+Chr(13)+Chr(10)
oMsg:Refresh()
oMsg:GoEnd() 
oDlgNFS:Refresh()
dDataBase := dDBaseBKP 
cEmpAnt := cEmpBKP  
cFilAnt := cFilBKP 
OpenFile(cEmpAnt+cFilAnt)
OpenSM0()
fLegFat("")
Return

//-------------------------------------------------------------------------
Static Function PPX_EXCNF(cPeriodo,nCboTipo,dDtImp,cDoPedido,cAtePedido,cSerieNF,p_par)
//-------------------------------------------------------------------------
Local aCabec   := {}
Local aItens   := {}
Local axLinha  := {}
Local aNotas   := {}
Local nX1      := 0
Local nX2      := 0
Local nQtdeNf  := 0
Local cMsgPV   := ""
Local cMsgNF   := ""

cMsg := "Notas Fiscais excluída(s)"+Chr(13)+Chr(10)
oMsg:Refresh()
oMsg:GoEnd() 
oDlgNFS:Refresh()
aArea := GetArea()
ProcRegua(Len(aWBrowse1))
For nX1 := 1 to Len(aWBrowse1)
    If aWBrowse1[nX1,1]
       Begin Transaction 
       aItens  := {}
       aCabec  := {}
       axLinha := {}
       dbSelectArea("PPX")
       PPX->(dbSetOrder(1))
       PPX->(dbSeek(xFilial("PPX")+aWBrowse1[nX1,3]))
       
       dbSelectArea("SC5")
       SC5->(dbSetOrder(1))
       SC5->(dbSeek(xFilial("SC5")+aWBrowse1[nX1,2]))
            
       dbSelectArea("SA1")
       SA1->(dbSetOrder(3))
       SA1->(dbSeek(xFilial("SA1")+PPX->PPX_CNPJ))
       
       dbSelectArea("SF2")
       SF2->(dbSetOrder(1))
       SF2->(dbSeek(xFilial("SF2")+Padr(SC5->C5_NOTA,TamSX3("F2_DOC")[1])+Padr(SC5->C5_SERIE,TamSX3("F2_SERIE")[1])+SC5->C5_CLIENTE+SC5->C5_LOJACLI+SPACE(TamSX3("F2_FORMUL")[1])+"N"))
       
       IncProc("Nota Fiscal "+Alltrim(SC5->C5_NOTA)+" - "+AllTrim(SC5->C5_SERIE)+"-"+ALLTRIM(SA1->A1_NREDUZ))
	      
       aItens := {}
	   aAdd(aCabec, {"F2_DOC"    ,SF2->F2_DOC     ,Nil})
	   aAdd(aCabec, {"F2_SERIE"  ,SF2->F2_SERIE   ,Nil})
	   lMsErroAuto := .F.
         MSExecAuto({|x| MATA520(x)},aCabec,5)   
       If (lMsErroAuto)
          MostraErro()
          DisarmTransaction()
          Return
       Endif
       aCabec := {}
       aItens := {}
       aAdd(aCabec,{"C5_FILIAL" ,xFilial("SC5") ,Nil})
	   aAdd(aCabec,{"C5_NUM"	,SC5->C5_NUM    ,Nil})
	   aAdd(aCabec,{"C5_TIPO"   ,SC5->C5_TIPO   ,Nil})
	   aAdd(aCabec,{"C5_CLIENTE",SC5->C5_CLIENTE,Nil})
	   aAdd(aCabec,{"C5_LOJACLI",SC5->C5_LOJACLI,Nil})
	   aAdd(aCabec,{"C5_CLIENT" ,SC5->C5_CLIENT ,Nil})
	   aAdd(aCabec,{"C5_LOJAENT",SC5->C5_LOJAENT,Nil})
	   aAdd(aCabec,{"C5_TIPOCLI",SC5->C5_TIPOCLI,Nil})
	   aAdd(aCabec,{"C5_EMISSAO",SC5->C5_EMISSAO,Nil})
	   aAdd(aCabec,{"C5_TIPLIB" ,SC5->C5_TIPLIB ,Nil})
	   aAdd(aCabec,{"C5_TPCARGA",SC5->C5_TPCARGA,Nil})
       dbSelectArea("SC6")
       SC6->(dbSetOrder(1))
       SC6->(dbSeek(xFilial("SC6")+SC5->C5_NUM))
       While !Eof() .And. SC6->C6_FILIAL == xFilial('SC6') .AND. SC6->C6_NUM = SC5->C5_NUM
             axLinha := {}
             aAdd(axLinha,{"C6_FILIAL" ,xFilial("SC6") ,Nil})
             aAdd(axLinha,{"C6_NUM"	   ,SC6->C6_NUM    ,Nil})
             aAdd(axLinha,{"C6_ITEM"   ,SC6->C6_ITEM   ,Nil})
             aAdd(axLinha,{"C6_PRODUTO",SC6->C6_PRODUTO,Nil})
             aAdd(axLinha,{"C6_QTDVEN" ,SC6->C6_QTDVEN ,Nil})
             aAdd(axLinha,{"C6_QTDLIB" ,SC6->C6_QTDLIB ,Nil})
             aAdd(axLinha,{"C6_UM"     ,SC6->C6_UM     ,Nil})      
             aAdd(axLinha,{"C6_PRCVEN" ,SC6->C6_PRCVEN ,Nil})
             aAdd(axLinha,{"C6_PRUNIT" ,SC6->C6_PRUNIT ,Nil})
             aAdd(axLinha,{"C6_VALOR"  ,SC6->C6_VALOR  ,Nil})
             aAdd(axLinha,{"C6_TPOP"   ,SC6->C6_TPOP   ,Nil})
             aAdd(axLinha,{"C6_SUGENTR",SC6->C6_SUGENTR,Nil})
             aAdd(axLinha,{"C6_VALDESC",SC6->C6_VALDESC,Nil})
             aAdd(axLinha,{"C6_TES"    ,SC6->C6_TES    ,Nil})
             aAdd(axLinha,{"C6_LOCAL"  ,SC6->C6_LOCAL  ,Nil})
             aAdd(axLinha,{"C6_CLI"    ,SC6->C6_CLI    ,Nil})
             aAdd(axLinha,{"C6_LOJA"   ,SC6->C6_LOJA   ,Nil})
             aAdd( aItens,axLinha )
             SC6->(dbSkip())
       Enddo
       lMsErroAuto := .F.
       MsExecAuto({|x,y,z| MATA410(x,y,z)},aCabec,aItens,4)
	   If lMsErroAuto
	      If File( cStartPath+ cArqLog )
		     Ferase( cStartPath+ cArqLog )
		 Endif
		 Mostraerro(cStartPath, cArqLog)
		 cMsg += ">>>>> Problemas na Alteração do Pedido : "+PPX->PPX_ID+" - "+SC5->C5_NUM+ Chr(13)+Chr(10)
		 If File( cStartPath+cArqLog )
		    cLog := MemoRead(cStartPath+cArqLog)
			cMsg += cLog+ Chr(13)+Chr(10)
		    Endif
		    MsgStop(cMsg,"Erro")
            Disarmtransaction()
            Return
       Endif   
       lMsErroAuto := .F.
       MsExecAuto({|x,y,z| MATA410(x,y,z)},aCabec,aItens,5)
	   If lMsErroAuto
	      If File( cStartPath+ cArqLog )
		     Ferase( cStartPath+ cArqLog )
		 Endif
		 Mostraerro(cStartPath, cArqLog)
		 cMsg += ">>>>> Problemas na Exclusão do Pedido : "+PPX->PPX_ID+" - "+SC5->C5_NUM+ Chr(13)+Chr(10)
		 If File( cStartPath+cArqLog )
		    cLog := MemoRead(cStartPath+cArqLog)
			cMsg += cLog+ Chr(13)+Chr(10)
		    Endif
		    MsgStop(cMsg,"Erro")
            Disarmtransaction()
            Return
       Endif  
       cMsgPV := ""
       cMsgNF := ""   
       If Select("QRY_AUX")>0
          DbSelectArea("QRY_AUX")
          DbCloseArea()
       Endif
       cQryAux := " SELECT * FROM "+RetSqlName("SC5")
       cQryAux += " WHERE D_E_L_E_T_ = ' ' "
       cQryAux += "AND C5_FILIAL = '"+xFilial("SC5")+"'"
       cQryAux += "AND LEFT(C5_OBS,6) = '"+PPX->PPX_ID+"' C*O*L*L*A*T*E*"
       cQryAux := If(cEmpAnt="99",Replace(cQryAux,"C*O*L*L*A*T*E*","COLLATE Latin1_General_CI_AS"),Replace(cQryAux,"C*O*L*L*A*T*E*",""))
       cQryAux := ChangeQuery(cQryAux)
       TCQuery cQryAux New Alias "QRY_AUX"
      // QRY_AUX->(DbGoTop())
      DbSelectArea("QRY_AUX")
       While !QRY_AUX->(Eof())	   
             dbSelectArea("SC5")
             SC5->(dbSetOrder(1))
             SC5->(dbSeek(xFilial("SC5")+ QRY_AUX->C5_NUM))
             If Alltrim(SC5->C5_NOTA) = "" 
                If PPX->PPX_AGLUT = "C" 
                   cMsgPV += "PV "+QRY_AUX->C5_NUM+" ==> Churrascarias "+Chr(13)+Chr(10)
                Endif    
                If PPX->PPX_AGLUT = "I" 
                   cMsgPV += "PV "+QRY_AUX->C5_NUM+" ==> Infrações"+Chr(13)+Chr(10)
                Endif    
                If PPX->PPX_AGLUT $ "|G|S|" 
                   If SC5->C5_XTPTAXA ="A"
                      cMsgPV += "PV "+QRY_AUX->C5_NUM+" ==> Administração"+Chr(13)+Chr(10)
                   Endif   
                   If SC5->C5_XTPTAXA ="M"
                      cMsgPV += "PV "+QRY_AUX->C5_NUM+" ==> Mídia"+Chr(13)+Chr(10)
                   Endif    
                   If SC5->C5_XTPTAXA ="S"
                      cMsgPV += "PV "+QRY_AUX->C5_NUM+" ==> Serviços"+Chr(13)+Chr(10)
                   Endif 
                Endif    
             Endif      
             If Alltrim(SC5->C5_NOTA) <> "" 
                If PPX->PPX_AGLUT = "C" 
                   cMsgNF += "==> NF "+Alltrim(SC5->C5_NOTA)+"-"+Alltrim(SC5->C5_SERIE)+" gerada para o PV 1"+QRY_AUX->C5_NUM+" ==> Churrascarias"+Chr(13)+Chr(10)
                Endif    
                If PPX->PPX_AGLUT = "I" 
                   cMsgNF += "==> NF "+Alltrim(SC5->C5_NOTA)+"-"+Alltrim(SC5->C5_SERIE)+" gerada para o PV 1"+QRY_AUX->C5_NUM+" ==> Infrações"+Chr(13)+Chr(10)
                Endif    
                If PPX->PPX_AGLUT $ "|G|S|" 
                   If SC5->C5_XTPTAXA ="A"
                      cMsgNF += "==> NF "+Alltrim(SC5->C5_NOTA)+"-"+Alltrim(SC5->C5_SERIE)+" gerada para o PV 1"+QRY_AUX->C5_NUM+" ==> Administração"+Chr(13)+Chr(10)
                   Endif   
                   If SC5->C5_XTPTAXA ="M"
                      cMsgNF += "==> NF "+Alltrim(SC5->C5_NOTA)+"-"+Alltrim(SC5->C5_SERIE)+" gerada para o PV 1"+QRY_AUX->C5_NUM+" ==> Mídia"+Chr(13)+Chr(10)
                   Endif    
                   If SC5->C5_XTPTAXA ="S"
                      cMsgNF += "==> NF "+Alltrim(SC5->C5_NOTA)+"-"+Alltrim(SC5->C5_SERIE)+" gerada para o PV 1"+QRY_AUX->C5_NUM+" ==> Serviços"+Chr(13)+Chr(10)
                   Endif 
                Endif    
             Endif                                                                                                                                                                                                                                                                                                             
       QRY_AUX->(dbSkip())
       Enddo
       RecLock("PPX",.F.)
       PPX->PPX_ERRO   := ""
       PPX->PPX_STATUS := '1'
       If Alltrim(cMsgPV) <> ""
          cMsgPV := "Pedidos de Venda gerados"+Chr(13)+Chr(10)+Alltrim(cMsgPV) 
       Endif    
       If Alltrim(cMsgNF) <> ""
           cMsgNF := +Chr(13)+Chr(10)+"Notas Fiscais geradas"+Chr(13)+Chr(10)+Alltrim(cMsgNF) 
           PPX->PPX_STATUS := '5'
          If PPX->PPX_AGLUT $ "|A||M|S|"
             PPX->PPX_STATUS := '4'
          Endif
       Endif    
       If Alltrim(cMsgNF) = ""  
          PPX->PPX_PEDIDO := ' '
       Endif
       PPX->PPX_ERRO   := Alltrim(cMsgPV)+Chr(13)+Chr(10)+Chr(13)+Chr(10)+Alltrim(cMsgNF)
       PPX->(msUnlock())
       fLegFat(PPX->PPX_ID)
       nQtdeNf ++
       cMsg += Time()+"--> Nota Fiscal excluídas "+Alltrim(SF2->F2_DOC)+"-"+Alltrim(SF2->F2_SERIE)+Chr(13)+Chr(10) 
       End Transaction        
    Endif       
    oMsg:Refresh()
    oMsg:GoEnd() 
    oDlgNFS:Refresh()
Next nX1 
cMsg += Time()+"--> "+Alltrim(Str(nQtdeNf))+" Documentos excluídos"+Chr(13)+Chr(10)
cMsg += Time()+"--> Finalizado a geração de Documentos de Saída (NF)"+Chr(13)+Chr(10)
oMsg:Refresh()
oMsg:GoEnd() 
oDlgNFS:Refresh()
oMsg:Refresh()
oMsg:GoEnd() 
oDlgNFS:Refresh()           
Return

//-------------------------------------------------------------------------
Static Function PPX_PVCOMPO()
//-------------------------------------------------------------------------
Local aHeaderEX := {}
Local aFields   := {"PPX_AGLUT","PPZ_EMPRES","PPZ_MASTER","PPZ_REDE","PPZ_CNPJ","PPZ_CLI","PPZ_LOJA","PPZ_LJREDE","PPZ_RAZAO","PPZ_FANTA","PPZ_ID","PPZ_PEDIDO","PPZ_ITEM","PPZ_EMISSA","PPZ_VENCTO","PPZ_SUGENT","PPZ_CODPRO","B1_DESC","PPZ_TES","PPZ_LOCAL","PPZ_QTDVEN","PPZ_PRCVEN","PPZ_VALOR","PPZ_VLDESC","PPZ_TXADM","PPZ_TXSRV","PPZ_TXMID"}
Local nX        := 0
Local aVetor    := {}
Local cPeriodo  := PPX->PPX_PERIOD
Local dDtImport:=  PPX->PPX_DATA
Local cNumId    := PPX->PPX_ID
Local aItens    := {}
Local aAglut    := {}
Local aGeral    := {}
Local aInfra    := {}
Local aChurr    := {}
Local aTodos    := {}
Local cQryAux   := ""
Local nTotal    := 0

dbSelectArea("SX3")
SX3->(DbSetOrder(2))
For nX := 1 to Len(aFields)
    If SX3->(DbSeek(aFields[nX]))
      aAdd(aHeaderEX,{SX3->X3_TITULO,;
                      SX3->X3_CAMPO,;
                      SX3->X3_PICTURE,;
                      SX3->X3_TAMANHO,;
                      SX3->X3_DECIMAL,;
                      "",;
                      "",;
                      SX3->X3_TIPO,;
                      SX3->X3_F3,;
                      SX3->X3_CONTEXT,;
                      SX3->X3_CBOX,;
                      SX3->X3_RELACAO})
    Endif                
 Next nX
 cQryAux   := ""
 If Select("QRY_AUX")>0
   DbSelectArea("QRY_AUX")
   DbCloseArea()
Endif
cQryAux += " SELECT PPX.*,PPZ.*, SB1.* FROM "+RetSqlName("PPX")+" PPX,"+RetSqlName("PPZ")+" PPZ,"+RetSqlName("SB1")+" SB1"
cQryAux += " WHERE PPX.D_E_L_E_T_ =  ' '"	 
cQryAux += " AND PPZ.D_E_L_E_T_ =  ' '"	 
cQryAux += " AND SB1.D_E_L_E_T_ =  ' '"	 
cQryAux += " AND PPX_FILIAL = '"+xFilial("PPX")+"'" 
cQryAux += " AND PPX_FILIAL = PPZ_FILIAL COLLATE Latin1_General_CI_AS" 
cQryAux += " AND B1_FILIAL = '"+xFilial("SB1")+"'" 
cQryAux += " AND PPX_ID = PPZ_ID COLLATE Latin1_General_CI_AS"
cQryAux += " AND PPZ_CODPRO = B1_COD COLLATE Latin1_General_CI_AS"
cQryAux += " AND PPX_PERIOD = '"+cPeriodo+"'" 
cQryAux += " AND PPX_DATA = '"+DTOS(dDtImport)+"'" 
cQryAux += " ORDER BY PPX_AGLUT,PPZ_EMPRES,PPZ_MASTER,PPZ_REDE,PPZ_RAZAO,PPX_ID"
cQryAux := ChangeQuery(cQryAux)
TCQuery cQryAux New Alias "QRY_AUX"
Count to nTotal
ProcRegua(nTotal)
QRY_AUX->(DbGoTop())
DbSelectArea("QRY_AUX")
While !QRY_AUX->(Eof())
       aItens   := {}
       For nX := 1 to Len(aFields)
           aAdd(aItens,QRY_AUX->&(aFields[nX]))
       Next nX
       If QRY_AUX->PPX_AGLUT = "S"
          aAdd(aAglut,aItens)
       Endif
       If QRY_AUX->PPX_AGLUT = "G"   
          aAdd(aGeral,aItens)
       Endif
       If QRY_AUX->PPX_AGLUT = "C"   
          aAdd(aChurr,aItens)
       Endif 
       If QRY_AUX->PPX_AGLUT = "I"   
          aAdd(aInfra,aItens)
       Endif 
       aAdd(aTodos,aItens)
      QRY_AUX->(DbSkip())      
EndDo    
If Len(aAglut) > 0
   aAdd(aVetor,{"Aglutinados","Composição de Pedidos de Venda - PPM(Aglutinados)",aHeaderEX,aAglut,"PPM_Composição"})
Endif   
If Len(aGeral) > 0
   aAdd(aVetor,{"Não Aglutinados","Composição de Pedidos de Venda - PPM(Não Aglutinados)",aHeaderEX,aGeral,"PPM_Composição"})
Endif 
If Len(aChurr) > 0
   aAdd(aVetor,{"Churrascarias","Composição de Pedidos de Venda - PPM(Churrascarias)",aHeaderEX,aChurr,"PPM_Composição"})
Endif 
If Len(aInfra) > 0
   aAdd(aVetor,{"Infrações","Composição de Pedidos de Venda - PPM(Infrações)",aHeaderEX,aInfra,"PPM_Composição"})
Endif 
If Len(aTodos) > 0
   aAdd(aVetor,{"Todos","Composição de Pedidos de Venda - PPM(Todos)",aHeaderEX,aTodos,"PPM_Composição"})
Endif 
If Len(aTodos) > 0
   U_VX_EXCEL(aVetor,.T.)
Endif
Return

//-------------------------------------------------------------------------
Static Function PPX_PLANFIN()
//-------------------------------------------------------------------------
Local aHeaderEX  := {}
Local aFields    := {"E1_CLIENTE","E1_LOJA","E1_NOMCLI","E1_PREFIXO","E1_NUM","E1_PARCELA","E1_TIPO","E1_EMISSAO","E1_VENCREA","E1_BAIXA","E1_VALOR"}
Local nX         := 0
Local aVetor     := {}
Local cPeriodo   := PPX->PPX_PERIOD
Local dDtImport  := PPX->PPX_DATA
Local dDtEmissao := PPX->PPX_EMISSA
Local cNumId     := PPX->PPX_ID
Local _cPerg     := 'HBFATPPM'
Local aItens     := {}
Local aAglut     := {}
Local aGeral     := {}
Local aInfra     := {}
Local aChurr     := {}
Local aTodos     := {}
Local cQryAux    := ""
Local nTotal     := 0

If !Pergunte(_cPerg,.T.)
   MsgAlert("Processamento cancelado com sucesso!","Algutinação de Títulos")
   Return Nil
 Endif
dbSelectArea("SX3")
SX3->(DbSetOrder(2))
For nX := 1 to Len(aFields)
    If SX3->(DbSeek(aFields[nX]))
      aAdd(aHeaderEX,{SX3->X3_TITULO,;
                      SX3->X3_CAMPO,;
                      SX3->X3_PICTURE,;
                      SX3->X3_TAMANHO,;
                      SX3->X3_DECIMAL,;
                      "",;
                      "",;
                      SX3->X3_TIPO,;
                      SX3->X3_F3,;
                      SX3->X3_CONTEXT,;
                      SX3->X3_CBOX,;
                      SX3->X3_RELACAO})
    Endif                
 Next nX
 cQryAux   := ""
 If Select("QRY_FT")>0
   DbSelectArea("QRY_FT")
   DbCloseArea()
Endif
cQryAux += " SELECT E1_CLIENTE,E1_LOJA,E1_NOMCLI,E1_PREFIXO,E1_NUM,E1_PARCELA,E1_TIPO,E1_EMISSAO,E1_VENCREA,E1_BAIXA,E1_VALOR "
cQryAux += " FROM "+RetSqlName("SE1")+" PAI"
cQryAux += " WHERE E1_TIPO = 'FT'"
cQryAux += " AND E1_FILIAL BETWEEN '"+MV_PAR01+"' AND '"+MV_PAR02+"'"
cQryAux += " AND E1_EMISSAO BETWEEN '"+DTOS(MV_PAR05)+"' AND '"+DTOS(MV_PAR06)+"'"
cQryAux += " AND D_E_L_E_T_ = ' '"
cQryAux += " ORDER BY E1_CLIENTE,E1_LOJA,E1_EMISSAO,E1_NUM"
cQryAux := ChangeQuery(cQryAux)
TCQuery cQryAux New Alias "QRY_FT"
Count to nTotal
ProcRegua(nTotal)
DbSelectArea("QRY_FT")
QRY_FT->(DbGoTop())
While !QRY_FT->(Eof())
       aItens   := {}
      For nX := 1 to Len(aFields)
           aAdd(aItens,If(aHeaderEX[nX,8]="D",STOD(QRY_FT->&(aFields[nX])),QRY_FT->&(aFields[nX])))
       Next nX
       aAdd(aTodos,aItens)
      If Select("QRYNF")>0
      DbSelectArea("QRYNF")
      DbCloseArea()
      Endif
       cQryAux   := ""
      cQryAux += " SELECT *  FROM "+RetSqlName("SE1")
      cQryAux += " WHERE E1_TIPO = 'NF'"
      cQryAux += " AND E1_FILIAL BETWEEN '"+MV_PAR01+"' AND '"+MV_PAR02+"'"
      cQryAux += " AND E1_EMISSAO = '"+QRY_FT->E1_EMISSAO+"'"
      cQryAux += " AND E1_CLIENTE = '"+QRY_FT->E1_CLIENTE+"'"
      cQryAux += " AND E1_LOJA = '"+QRY_FT->E1_LOJA+"'"
      cQryAux += " AND D_E_L_E_T_ = ' '"
      cQryAux += " ORDER BY E1_CLIENTE,E1_LOJA,E1_EMISSAO,E1_NUM"
      cQryAux := ChangeQuery(cQryAux)
      TCQuery cQryAux New Alias "QRYNF"
      DbSelectArea("QRYNF")
      QRYNF->(DbGoTop())
      While !QRYNF->(Eof())
            aItens   := {}
      For nX := 1 to Len(aFields)
           aAdd(aItens,If(aHeaderEX[nX,8]="D",STOD(QRYNF->&(aFields[nX])),QRYNF->&(aFields[nX])))
       Next nX
            aAdd(aTodos,aItens)
            QRYNF->(DbSkip())
       Enddo
       aItens   := {}
       For nX := 1 to Len(aFields)
           aAdd(aItens," ")
       Next nX
       aAdd(aTodos,aItens)
      QRY_FT->(DbSkip())      
EndDo    
If Len(aTodos) > 0
   aAdd(aVetor,{"Titulos","Titulos a Receber-Aglutinados",aHeaderEX,aTodos,"PPM_Aglutinação"})
   U_VX_EXCEL(aVetor,.F.)
Endif
Return

//-------------------------------------------------------------------------
Static Function PPX_PLANNF()
//-------------------------------------------------------------------------
Local aHeaderEX  := {}
//Local aFields  := {"A1_COD","A1_LOJA","A1_NOME","A1_CGC","A1_EST","A1_MUN","A1_END","PPX_ID","PPX_PERIOD","PPX_DATA","PPX_AGLUT","C5_XTPTAXA","F2_DOC","F2_SERIE","F2_EMISSAO","F2_VALMERC","F2_VALICM","F2_VALIPI","F2_VALISS","F2_VALCOFI","F2_VALPIS","F2_VALBRUT","C5_NUM","C5_EMISSAO","B1_COD","B1_DESC","D2_TES","D2_QUANT","D2_PRCVEN","D2_TOTAL","D2_VALICM","D2_VALIPI","D2_VALISS","D2_VALCOF","D2_VALPIS"}
Local aFields    := {"A1_COD","A1_LOJA","A1_NOME","A1_CGC","A1_EST","A1_MUN","PPX_ID","PPX_PERIOD","PPX_DATA","PPX_AGLUT","C5_XTPTAXA","F2_DOC","F2_SERIE","F2_EMISSAO","C5_NUM","C5_EMISSAO","B1_COD","B1_DESC","D2_TES","D2_PRCVEN","D2_TOTAL","D2_VALICM","D2_VALISS","D2_VALCOF","D2_VALPIS"}
Local aTitulos   := {"Cliente","Loja","Razão Social","CNPJ/CPF","UF","Município","Id Imp","Período","Data Imp","Tp Aglut","Tp Taxa","Número NF","Série","Emissão NF","Núm PV","Emissão PV","Produto","Descrição","TES","Pr Unit","Vr Total","Vr ICMS","Vr ISS","Vr Cofins","Vr Pis"}
Local nX         := 0
Local aVetor     := {}
Local cPeriodo   := PPX->PPX_PERIOD
Local dDtImport  :=  PPX->PPX_DATA
Local dDtEmissao :=  PPX->PPX_EMISSA
Local cNumId     := PPX->PPX_ID
Local _cPerg     := 'HBFATPPM'
Local aItens     := {}
Local aAglut     := {}
Local aGeral     := {}
Local aInfra     := {}
Local aChurr     := {}
Local aTodos     := {}
Local cQryAux    := ""
Local nTotal     := 0
Local cQuebra    := ""

dbSelectArea("SX3")
SX3->(DbSetOrder(2))
For nX := 1 to Len(aFields)
    If SX3->(DbSeek(aFields[nX]))
      aAdd(aHeaderEX,{aTitulos[nX],;
                      SX3->X3_CAMPO,;
                      SX3->X3_PICTURE,;
                      SX3->X3_TAMANHO,;
                      SX3->X3_DECIMAL,;
                      "",;
                      "",;
                      SX3->X3_TIPO,;
                      SX3->X3_F3,;
                      SX3->X3_CONTEXT,;
                      SX3->X3_CBOX,;
                      SX3->X3_RELACAO})
    Endif                
 Next nX
 cQryAux   := ""
 If Select("QRY_FT")>0
   DbSelectArea("QRY_FT")
   DbCloseArea()
Endif
cQryAux   += " SELECT A1_COD,A1_LOJA,A1_NOME,A1_CGC,A1_EST,A1_MUN,A1_END,PPX_ID,PPX_PERIOD,PPX_DATA"
cQryAux   += " ,CASE PPX_AGLUT WHEN 'I' THEN 'INFRAÇÃO'"
cQryAux   += " 				WHEN 'C' THEN 'CHURRASCARIAS'"
cQryAux   += " 				WHEN 'G' THEN 'GERAL'"
cQryAux   += " 				WHEN 'S' THEN 'AGLUTINADOS'"
cQryAux   += " 				ELSE 'NORMAL' END PPX_AGLUT"
cQryAux   += " ,CASE C5_XTPTAXA WHEN 'A' THEN 'ADMINISTRAÇÃO'"
cQryAux   += "                  WHEN 'M' THEN 'MÍDIA'"
cQryAux   += " 				 WHEN 'I' THEN 'INFRAÇÃO'"
cQryAux   += " 				 WHEN 'C' THEN 'CHURRASCARIAS'"
cQryAux   += " 				 WHEN 'S' THEN 'SERVIÇOS'"
cQryAux   += " 				 ELSE 'NORMAL' END C5_XTPTAXA"
cQryAux   += ",F2_DOC,F2_SERIE,F2_EMISSAO,F2_VALMERC,F2_VALICM,F2_VALIPI,F2_VALISS,F2_VALCOFI"
cQryAux   += ",F2_VALPIS,F2_VALBRUT,C5_NUM,C5_EMISSAO,B1_COD,B1_DESC,D2_TES,D2_QUANT,D2_PRCVEN,D2_TOTAL"
cQryAux   += ",D2_VALICM,D2_VALIPI,D2_VALISS,D2_VALCOF,D2_VALPIS"
cQryAux   += " FROM" +RetSqlName("PPX")+" PX"
cQryAux   += " INNER JOIN " +RetSqlName("SC5")+" AS C5 ON C5.D_E_L_E_T_ = ' ' "
cQryAux   += "                                         AND C5_FILIAL = '"+xFilial("SC5")+"'"
cQryAux   += "                                         AND LEFT(C5_OBS,6) = PPX_ID" 
cQryAux   += " INNER JOIN " +RetSqlName("SC6")+" AS C6 ON C6.D_E_L_E_T_ = ' '" 
cQryAux   += "                         AND C6_FILIAL = C5_FILIAL" 
cQryAux   += " 						AND C6_NUM = C5_NUM"
cQryAux   += " 						AND C6_CLI = C5_CLIENTE"
cQryAux   += " 						AND C6_LOJA = C5_LOJACLI"
cQryAux   += " INNER JOIN "+RetSqlName("SF2")+" AS F2 ON F2.D_E_L_E_T_ = ' ' 
cQryAux   += "                                         AND F2_FILIAL = C5_FILIAL"
cQryAux   += " 						                   AND F2_DOC = C5_NOTA"
cQryAux   += " 						                   AND F2_SERIE = C5_SERIE"
cQryAux   += " INNER JOIN "+RetSqlName("SD2")+" AS D2 ON D2.D_E_L_E_T_ = ' '" 
cQryAux   += "                                         AND D2_FILIAL = F2_FILIAL" 
cQryAux   += " 						                   AND F2_DOC = D2_DOC"
cQryAux   += " 						                   AND D2_SERIE = F2_SERIE"
cQryAux   += " 						                   AND D2_CLIENTE = F2_CLIENTE"
cQryAux   += " 						                   AND D2_LOJA = F2_LOJA"
cQryAux   += " INNER JOIN "+RetSqlName("SA1")+" AS A1 ON A1.D_E_L_E_T_ = ' '" 
cQryAux   += "                                         AND A1_FILIAL = '"+xFilial("SA1")+"'"
cQryAux   += "                                         AND A1_COD = C5_CLIENTE" 
cQryAux   += " 						                   AND A1_LOJA = C5_LOJACLI"
cQryAux   += " INNER JOIN "+RetSqlName("SB1")+"  AS B1 ON B1.D_E_L_E_T_ = ' '" 
cQryAux   += "                                          AND B1_FILIAL = '"+xFilial("SB1")+"'"
cQryAux   += "                                         AND B1_COD = D2_COD"
cQryAux   += " WHERE PX.D_E_L_E_T_ = ' '"
cQryAux   += " AND PPX_FILIAL = '"+xFilial("PPX")+"'"
cQryAux   += " ORDER BY A1_COD,A1_LOJA,F2_DOC,F2_SERIE,C5_NUM"
cQryAux := ChangeQuery(cQryAux)
TCQuery cQryAux New Alias "QRY_FT"
Count to nTotal
ProcRegua(nTotal)
DbSelectArea("QRY_FT")
QRY_FT->(DbGoTop())
While !QRY_FT->(Eof())
       aItens   := {}
       If Alltrim(cQuebra) <> Alltrim(QRY_FT->A1_COD)+Alltrim(QRY_FT->A1_LOJA)+Alltrim(QRY_FT->F2_DOC)+Alltrim(QRY_FT->F2_SERIE)+Alltrim(QRY_FT->C5_NUM)
          If Alltrim(cQuebra) <> "" 
             aItens := {} 
             For nX := 1 to Len(aFields)
                  aAdd(aItens,If(aHeaderEX[nX,8]="D",CTOD("  /  /    "),If(aHeaderEX[nX,8]="N",0,"")))  
             Next nX
             aAdd(aTodos,aItens)
          Endif   
          aItens := {}
          For nX := 1 to Len(aFields)
             If Left(aFields[nX],3) = "PPX" .OR. Left(aFields[nX],2) $ "|C5|F2|A1|"
                aAdd(aItens,If(aHeaderEX[nX,8]="D",STOD(QRY_FT->&(aFields[nX])),QRY_FT->&(aFields[nX])))
             Else  
                aAdd(aItens,If(aHeaderEX[nX,8]="D",CTOD("  /  /    "),If(aHeaderEX[nX,8]="N",0,"")))
             Endif        
          Next nX
          aAdd(aTodos,aItens)
       Endif   
       aItens   := {}
       For nX := 1 to Len(aFields)
           If Left(aFields[nX],3) = "PPX" .OR. Left(aFields[nX],2) $ "|C5|F2|A1|"
              aAdd(aItens,If(aHeaderEX[nX,8]="D",CTOD("  /  /    "),If(aHeaderEX[nX,8]="N",0,"")))
         Else 
              aAdd(aItens,If(aHeaderEX[nX,8]="D",STOD(QRY_FT->&(aFields[nX])),QRY_FT->&(aFields[nX])))
         Endif     
       Next nX
       aAdd(aTodos,aItens)
       cQuebra := Alltrim(QRY_FT->A1_COD)+Alltrim(QRY_FT->A1_LOJA)+Alltrim(QRY_FT->F2_DOC)+Alltrim(QRY_FT->F2_SERIE)+Alltrim(QRY_FT->C5_NUM)
       QRY_FT->(DbSkip())      
EndDo    
If Len(aTodos) > 0
   aAdd(aVetor,{"Movimentos",Alltrim(SM0->M0_FILIAL)+" - Histórico de Importação ( Período: "+Alltrim(PPX->PPX_PERIOD)+" )",aHeaderEX,aTodos,"xMovimentos"})
   U_VX_EXCEL(aVetor,.F.)
Endif
Return

//-------------------------------------------------------------------------
Static Function fLerEmpresa(p_Prefix)
//-------------------------------------------------------------------------
Local cQuery   := ""
Local cTitulo  := ""
Local cPrefix  := ""
Local cEmpresx  := ""
Local axTitulo  := {}

If Select("TMPEMP")>0
   DbSelectArea("TMPEMP")
   DbCloseArea()
Endif
cQuery := " SELECT * FROM " + RetSqlName('SX5')
cQuery += " WHERE D_E_L_E_T_ = ' '  "
cQuery += " AND X5_TABELA  = 'Z3'"
cQuery += " AND X5_CHAVE = '"+p_Prefix+"'"
TCQuery cQryAux New Alias "TMPEMP"
If !TMPEMP->(Eof()) .and. !TMPEMP->(Bof())
	cTitulo := Alltrim(TMPEMP->X5_DESCRI)+"-"+Alltrim(TMPEMP->X5_CHAVE)+"-"+Alltrim(TMPEMP->X5_DESCSPA)
	cPrefix := Alltrim(TMPEMP->X5_CHAVE)
	cEmpresx := Alltrim(TMPEMP->X5_DESCRI)
	TMPEMP->(dbSkip())
Endif
axTitulo := {cTitulo,cPrefix,cEmpresx}
Return (axTitulo)

//-------------------------------------------------------------------------
Static Function fLerSX5(cTabela,cDescri)
//-------------------------------------------------------------------------
Local cQuery := ""

If Select("TMPSX5")>0
   DbSelectArea("TMPSX5")
   DbCloseArea()
Endif
cQuery := " SELECT * FROM " + RetSqlName('SX5')
cQuery += " WHERE D_E_L_E_T_ = ' '  "
cQuery += " AND X5_TABELA  = '"+cTabela+"'"
cQuery += " AND UPPER(X5_DESCRI) = '"+UPPER(AllTrim(cDescri))+"'"
TCQuery cQuery New Alias "TMPSX5"
If !TMPSX5->(Eof()) .and. !TMPSX5->(Bof())
	cQuery := Alltrim(TMPSX5->X5_CHAVE)
Else
    cQuery := UPPER(AllTrim(cDescri))
Endif
Return (cQuery)

//-------------------------------------------------------------------------
Static Function fLerSA1(p_CNPJ)
//-------------------------------------------------------------------------
Local cCnpj  := p_CNPJ
Local nRecno := 0

If Select("TMPSA1")>0
   DbSelectArea("TMPSA1")
   DbCloseArea()
Endif
cQryAux := "SELECT A1_CGC,R_E_C_N_O_ A1_RECNO,A1_MSBLQL FROM " + RetSqlName('SA1')
cQryAux += " WHERE D_E_L_E_T_ = ' '  "
cQryAux += " AND A1_FILIAL = '"+xFilial("SA1")+"'"
cQryAux += " AND A1_CGC LIKE '%"+ALLTRIM(cCnpj)+"%'"
cQryAux += "  ORDER BY A1_MSBLQL"
TCQuery cQryAux New Alias "TMPSA1"	
While !TMPSA1->( Eof() )
    nRecno := TMPSA1->A1_RECNO
	cCnpj := TMPSA1->A1_CGC
	TMPSA1->(dbSkip())
Enddo
Return ({cCnpj,nRecno})

//-------------------------------------------------------------------------
Static Function ZRetLog(aErr,cLit)
//-------------------------------------------------------------------------

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

//-------------------------------------------------------------------------
Static Function PPX_TRANSMNF(cPeriodo,nCboTipo,dDtImp,cDoPedido,cAtePedido,cSerieNF,p_par)
//-------------------------------------------------------------------------
If IW_MsgBox("Deseja efetuar a Transmissão da Nota Fiscal [ "+ALLTRIM(SF2->F2_DOC)+" / "+ALLTRIM(SF2->F2_SERIE)+" ] ?","Atencao","YESNO")
   nX2 := 0
   cMsg := "Notas Fiscais Transmitida(s)"+Chr(13)+Chr(10)
   oMsg:Refresh()
   oMsg:GoEnd() 
   oDlgNFS:Refresh()
   aArea := GetArea()
   ProcRegua(Len(aWBrowse1))
   For nX1 := 1 to Len(aWBrowse1)
       If aWBrowse1[nX1,1] 
          dbSelectArea("SC5")
          SC5->(dbSetOrder(1))
          If SC5->(dbSeek(xFilial("SC5")+aWBrowse1[nX1,2]))
             dbSelectArea("SF2")
             SF2->(dbSetOrder(1))
             If SF2->(dbSeek(xFilial("SF2")+SC5->C5_CLIENTE+SC5->C5_LOJACLI+SC5->C5_NOTA+SC5->C5_SERIE))
                U_EnvSefaz(SF2->F2_SERIE,SF2->F2_DOC)
                cMsg := "NF: "+SF2->F2_SERIE+"/"+SF2->F2_DOC+" transmitida"+Chr(13)+Chr(10)
                nX2 ++
                oMsg:Refresh()
                oMsg:GoEnd() 
                oDlgNFS:Refresh()  
             Endif   
          Endif   
       Endif   
   Next nX1   
Endif
cMsg := Alltrim(Str(nX2))+"NFs transmitida(s) "+Chr(13)+Chr(10)
cMsg := "Favor verificar no Monitor."+Chr(13)+Chr(10)
nX2 ++
oMsg:Refresh()
oMsg:GoEnd() 
oDlgNFS:Refresh()
Return

//-------------------------------------------------------------------------
Function U_EnvSefazS(cSerie,cDoc)
//-------------------------------------------------------------------------
Local cURL     := ""
Local lOk      := .T. 
Local oWs
Local cAmbiente

oWs     := WsSpedCfgNFe():New()
cURL     := PADR(GetMv("MV_SPEDURL"),250)
If CTIsReady()
   oWS:cUSERTOKEN := "TOTVS"
   oWS:cID_ENT    := cIdEnt
   oWS:nAmbiente := 0
   oWS:_URL       := AllTrim(cURL)+"/SPEDCFGNFe.apw"
   lOk := oWS:CFGAMBIENTE()
   cAmbiente := oWS:cCfgAmbienteResult
   cAmbiente := Substr(cAmbiente,1,1)
   AutoNfeEnv(cEmpAnt,cEmpAnt,"0",cAmbiente,cSerie,cDoc,cDoc)
Endif
Return

//-------------------------------------------------------------------------
Static Function fUltimaData()
//-------------------------------------------------------------------------
Local dXUltData := CTOD("//")

If Select("QRYNFULT")>0
   DbSelectArea("QRYNFULT")
   DbCloseArea()
Endif
cQryAux := " SELECT MAX(F2_EMISSAO) F2_EMISSAO FROM "+RetSqlName("SF2")
cQryAux += " WHERE F2_FILIAL = '"+xFilial("SF2")+"'"
cQryAux += " AND D_E_L_E_T_ = ' ' "
cQryAux := ChangeQuery(cQryAux)
TCQuery cQryAux New Alias "QRYNFULT"
dbselectarea("QRYNFULT")                   
QRYNFULT->(DbGoTop())
While QRYNFULT->(!EOF())  
      dXUltData := STOD(QRYNFULT->F2_EMISSAO)
      QRYNFULT->(dbSkip())
Enddo
//ALERT(dXUltData)
Return (dXUltData)

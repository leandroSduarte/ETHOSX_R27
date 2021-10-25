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
//=============================================================
User Function ALLIMPFIN()
//=============================================================
Private axArea      := GetArea()
Private oBrowse     := FwLoadBrw("ALLIMPFIN")
Private cCadastro   := "Importação de Títulos a Receber"
Private lMsErroAuto := .T.
Private lMarcar  	:= .F.
Private axLinha     := {}
Private bKeyF12	    
Private cEmpX       := ""
Private cFilX       := ""
Private aTitulo     := {}
Private aProcArq    := {}
Private aProcPed    := {}
Private cQryAux     := ""
Private oMark
Private cMarca      := ""
Private lInverte    := .F.
Private aCNPJ       := {}
Private nRecno      := 0
Private cLog        := ""
Private oMsg
Private cMsg := ""
Private cEmpBKP     := cEmpAnt //cEmpBKP+cFilBKP
Private cFilBKP     := cFilAnt
Private aRotina     := MenuDef()
Private cPrefxTel   := ""
Private cQryAux     := ""

cLog := "DELETE "+RetSqlName("PA2")+" WHERE D_E_L_E_T_ <> ' '"
TCSQLExec(cLog)
cLog := "UPDATE "+RetSqlName("PA2")+" SET PA2_OK = '  '"
TCSQLExec(cLog)
cPrefxTel := U_ETX_PREFIXO()

dbSelectArea("PA1")
dbSelectArea("PA2")
dbSelectArea("PPZ")

oMark := FWMarkBrowse():New()
cMarca      := oMark:Mark()
lInverte    := oMark:IsInvert()

bKeyF12	:= {||  oMark:SetInvert(.F.),oMark:Refresh(),oMark:GoTop(.T.) } //Programar a tecla F12
oMark:SetAlias('PA2')
oMark:SetFieldMark( 'PA2_OK' )
oMark:SetAllMark( { || oMark:AllMark() } )
oMark:SetDescription('Importação de Títulos a Receber')
oMark:AddLegend("PA2_STATUS=='0'","BR_VERMELHO","Importado com Erro" )
oMark:AddLegend("PA2_STATUS=='1'","BR_VERDE" ,"Importado com sucesso" )
oMark:AddLegend("PA2_STATUS=='2'","BR_AZUL"    ,"Títulos a Receber gerado")
oMark:bAllMark := { || U_PA2_INVERT (oMark:Mark(),lMarcar := !lMarcar ), oMark:Refresh(.T.)  }
oMark:Activate()
Return

//=======================================================================================================================================
User Function PA2_INVERT ()
//=======================================================================================================================================
cMarca    := oMark:Mark()
lInverte  := oMark:IsInvert()
cPrefxTel := U_ETX_PREFIXO()

dbSelectArea("PA2")
PA2->(dbGoTop())
While !PA2->( Eof() )
      If PA2->PA2_STATUS = "1"
         RecLock("PA2",.F.)
         PA2->PA2_OK := IIf(lMarcar,cMarca,'  ')
         PA2->(MsUnlock())
      Endif   
      PA2->(dbSkip())
EndDo
Return

//=============================================================
Static Function BrowseDef()
//=============================================================
Local oBrowse := FwMBrowse():New()

oBrowse:SetAlias("PA2")
oBrowse:SetDescription("Importação de Títulos a Receber")
oBrowse:SetMenuDef("ALLIMPFIN")
Return (oBrowse)

//=======================================================================================================================================
Static Function MenuDef()
//=======================================================================================================================================
Local aRotina  := {}  //FwMVCMenu("ALLIMPFIN")     
Local aRotina2 := {}
Local aRotina3 := {}
Local aRotina4 := {} 
Local aRotina5 := {}
Local cPrefxTel := ""

aAdd(aRotina,{"Visualizar","VIEWDEF.ALLIMPFIN",0,02})  
aAdd(aRotina,{"Incluir"   ,"VIEWDEF.ALLIMPFIN",0,03})  
aAdd(aRotina,{"Alterar"   ,"VIEWDEF.ALLIMPFIN",0,04})  
aAdd(aRotina,{"Excluir"   ,"VIEWDEF.ALLIMPFIN",0,05})  

aAdd(aRotina2,{"Importar"         ,"U_ALL_EXFIN('Importar')"      ,0,10})  
aAdd(aRotina2,{"Avaliar"          ,"U_ALL_EXFIN('Avaliar')"       ,0,11})   
aAdd(aRotina2,{"Excl Imp em Lote" ,"U_ALL_EXFIN('LoteReg')"       ,0,12})    
 

aAdd(aRotina3,{"Gerar Títulos"  ,"U_ALL_EXFIN('Gerar')"           ,0,14})  
aAdd(aRotina3,{"Excluir Títulos","U_ALL_EXFIN('LoteTit')"         ,0,15}) 

aAdd(aRotina,{"Importação"    ,aRotina2                           ,0,09})  
aAdd(aRotina,{"Títulos"       ,aRotina3                           ,0,13}) 

aAdd(aRotina,{"Documento"    ,"U_ALL_EXFIN('Documento')"          ,0,25})  
aAdd(aRotina,{"Legenda"   ,"U_ALL_LEGFIN()"                       ,0,26})
Return (aRotina)


//=============================================================
Static Function ModelDef()
//=============================================================
Local oModel := MPFormModel():New("PAOLLAM",,)
Local oStruPA2 := FwFormStruct(1, "PA2")

oModel:AddFields("PA2MASTER", NIL, oStruPA2)
oModel:SetPrimaryKey({'PA2_FILIAL'})
oModel:SetDescription("Importação de Títulos a Receber" )
oModel:GetModel("PA2MASTER"):SetDescription("Importação de Títulos a Receber")
Return (oModel)

//=============================================================
Static Function ViewDef()
//=============================================================
Local nXtamX := 100
Local oView := FwFormView():New()
Local oStruPA2 := FwFormStruct(2, "PA2")
Local oModel := FwLoadModel("ALLIMPFIN")

oView:SetModel(oModel)
oView:AddField("VIEW_PA2", oStruPA2, "PA2MASTER")
oView:CreateHorizontalBox("SUPERIOR", nXtamX)
oView:SetOwnerView("VIEW_PA2", "SUPERIOR")
Return (oView)

//=================================================================================================================
User Function ALL_LEGFIN()
//=================================================================================================================
Local aLegenda := {}

aAdd(aLegenda,{"BR_VERMELHO","Importação com Erro" })
aAdd(aLegenda,{"BR_VERDE" ,"Importação com sucesso" })
aAdd(aLegenda,{"BR_AZUL"    ,"Títulos a Receber gerado" })
BrwLegenda( cCadastro, "Legenda", aLegenda )
Return Nil

//==============================================================================
Static Function PA2_PRETELA(oMdl)
//==============================================================================
Local nX := 0
Return (.T.)

//==============================================================================
Static Function PA2_POSTELA(oMdl)
//==============================================================================
Local nX := 0
Return (.T.)

//==============================================================================
User Function ALL_EXFIN(p_Funcao)
//==============================================================================
Local X := 0
Local Xy := 0
 
Do Case
      Case Upper(Alltrim(p_Funcao)) = "IMPORTAR"
           Processa({|| PA2_IMPORT()},"Importando Planilha Excel")
    Case Upper(Alltrim(p_Funcao)) = "AVALIAR"
           Processa({|| PA2_AVALIAR()},"Analisando Títulos a Receber com Erro")           
      Case Upper(Alltrim(p_Funcao)) = "GERAR"
           Processa({|| PA2_GERA(1)},"Gerando Títulos a Receber")
      Case Upper(Alltrim(p_Funcao)) = "LOTEREG"
           PA2_TELOTE()
      Case Upper(Alltrim(p_Funcao)) = "DOCUMENTO"
           Processa({|| PA2_DOCUMENTO()},"Visualizar Documento")
      Case Upper(Alltrim(p_Funcao)) = "LOTETIT"
           Processa({|| PA2_EXCTIT()},"Exclusão de Títulos Gerados")
EndCase  
cLog := "UPDATE "+RetSqlName("PA2")+" SET PA2_OK = '  '"
TCSQLExec(cLog)     
Return

//=============================================================================================================================
Static Function PA2_IMPORT()
//=============================================================================================================================   
Local aCampos   := {}  
Local aDados    := {}           
Local aFields   := {"PA2_PERIOD","PA2_EMPRES","PA2_EMISSA","PA2_VALOR","PA2_DESC","PA2_VENCTO","PA2_CNPJ","PA2_FANTAS","PA2_RAZAO","PA2_MASTER","PA2_LOJA","PA2_REDE","PA2_PRODUT"}
Local aCabec    := {"Periodo"   ,"Empresa"   ,"Dt Emissao" ,"Valor"    ,"Desconto","Vencimento","CNPJ"    ,"Fantasia"  ,"Razão"    ,"Master"    ,"Loja"    ,"Rede"    ,"Protheus"}
Local aMeses    := {"JAN","FEV","MAR","ABR","MAI","JUN","JUL","AGO","SET","OUT","NOV","DEZ"}
Local astru     :={}
Local cAno      := 0
Local cAnomes   := ""
Local cArqMacro := "XLS2DBF.XLA" 
Local cPatch	:= ""//Alltrim(cDirCSV)+alltrim(GetNewPar('US_CSVDIR','ALIMPFIN\')) 
Local cDirCSV	:= GetSrvProfString("Startpath","")  
Local cEmp      := ""//FWCodEmp()
Local cFil      := ""//FWCodFil()
Local cFileC    := ""//"ALLTRIM(cDirCSV)+ALLTRIM(cFile)
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
Local lGera     := .F.
Local lFirst    := .T.
Local cCNPJ     := ""
Local cPagto    := ""
Local cNatFin   := ""
Local cQuebra   := "ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ"
Local dEmissao  := CTOD("  /  /    ")
Local nValor    := 0
Local nDesconto := 0
Local nXtam     := 0
Local RelEmp     := ""  //VX"   //GetMV("US_PREF")
Local cPathTmp    := cGetFile( '', 'Selecione Diretório onde estao os arquivos a serem processados', 0, , .F.,   GETF_LOCALHARD +  GETF_RETDIRECTORY + GETF_NETWORKDRIVE)

aFields := {}
aAdd(aFields,{"PA2_PERIOD","Periodo"     ,"",0,0,"UPPER(ALLTRIM(axLinha[01]))"})
aAdd(aFields,{"PA2_EMPRES","Empresa"     ,"",0,0,"LEFT(ALLTRIM(axLinha[02])+SPACE(500),TamSx3('PA2_EMPRES')[1])"})
aAdd(aFields,{"PA2_EMISSA","Dt Emissão"  ,"",0,0,"If(Valtype(axLinha[03])='D',axLinha[03],CTOD(axLinha[03]))"})
aAdd(aFields,{"PA2_VALOR" ,"Valor"       ,"",0,0,"Val(Replace(Replace(axLinha[04],'.',''),',','.'))"})
aAdd(aFields,{"PA2_DESC"  ,"Descontos"   ,"",0,0,"Val(Replace(Replace(axLinha[05],'.',''),',','.'))"})
aAdd(aFields,{"PA2_VENCTO","Dt Vencto"   ,"",0,0,"If(Valtype(axLinha[06])='D',axLinha[06],CTOD(axLinha[06]))"})
aAdd(aFields,{"PA2_CNPJ"  ,"CNPJ"        ,"",0,0,"REPLACE(REPLACE(REPLACE(axLinha[07],'.',''),'-',''),'/','')"})
aAdd(aFields,{"PA2_FANTAS","Fantasia"    ,"",0,0,"LEFT(ALLTRIM(axLinha[08])+SPACE(500),TamSx3('PA2_FANTAS')[1])"})
aAdd(aFields,{"PA2_RAZAO" ,"Razão Social","",0,0,"LEFT(ALLTRIM(axLinha[09])+SPACE(500),TamSx3('PA2_RAZAO')[1])"})
aAdd(aFields,{"PA2_MASTER","Master"      ,"",0,0,"LEFT(ALLTRIM(axLinha[10])+SPACE(500),TamSx3('PA2_MASTER')[1])"})
aAdd(aFields,{"PA2_LOJA"  ,"Loja"        ,"",0,0,"STRZERO(VAL(axLinha[11]),TamSx3('PA2_LOJA')[1])"})
aAdd(aFields,{"PA2_REDE"  ,"Rede"        ,"",0,0,"LEFT(ALLTRIM(axLinha[12])+SPACE(500),TamSx3('PA2_REDE')[1])"})
aAdd(aFields,{"PA2_PRODUT","Produto"     ,"",0,0,"LEFT(ALLTRIM(axLinha[13])+SPACE(500),TamSx3('PA2_PRODUT')[1])"})

For nX1 := 1 to Len(aFields)
    dbSelectArea("SX3")
    SX3->(DbSetOrder(2))
    If SX3->(DbSeek(aFields[nX1,1]))
       aFields[nX1,3] := SX3->X3_TIPO
       aFields[nX1,4] := SX3->X3_TAMANHO
       aFields[nX1,5] := SX3->X3_DECIMAL
    Endif
Next nX
aArq    := directory(Alltrim(cPathTmp)+"*.csv") 
If Len(aArq) < 1
   Alert("Não foram encontrados arquivos com Extensão [CSV]")
   Return
Endif   
For nX1 := 1 to Len(aArq)
    cFileC := Alltrim(cPathTmp)+Replace(Alltrim(aArq[nX1,1]),"\\","\")
    nHandle  := Ft_Fuse(cFileC)
    If nHandle == -1
       Help(,,"Help","Importação de Títulos a Receber", "Arquivo não existe ou está em uso", 1, 0)
       Return 
    Endif 
    aTitulo := U_ETX_EMPRE(Left(aArq[nX1,1],2))
    If Alltrim(aTitulo[1]) = ""
       Alert("Empresa não cadastrada na tabela SX5")
       Return
    Endif  
    dbCloseAll()
    cEmpX := aTitulo[2]
    cFilX := aTitulo[3]
    cEmpAnt := cEmpX
    cFilAnt := cFilX
    OpenFile(cEmpAnt+cFilAnt)
    OpenSM0()
    nLin    := 0
    nLinTit := 1
    nLinTot := 0
    aDados  := {}  
    axLinha := {}
    cLinha  := ""
    lFirst  := .T.  
	cQuebra := "ZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZZ"
	Ft_FGoTop()                                                         
	nLinTot := FT_FLastRec()-1
	ProcRegua(nLinTot)
	While nLinTit > 0 .AND. !Ft_FEof() //Pula as linhas de cabeçalho
	      Ft_FSkip()
	      nLinTit--
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
	      nXtam   := Len(aFields)
	      For nX3 := 1 to nXtam
	          axLinha[nX3]:= &(aFields[nX3,6]) 
	          nX3 := nX3
	      Next nX3   
	      aAdd(aDados,axLinha)
	      FT_FSkip()
	EndDo 
	nLin  := 0
	For nX2 := 1 to LEN(aDados)
	    cNumero := GetSxeNum('PA2','PA2_ID')
	    ConfirmSX8()
	    cMsg := ""
	    dbSelectArea("PA2") 
	    Reclock("PA2",.T.) 
	    PA2->PA2_FILIAL := XFILIAL("PA2")
	    PA2->PA2_ID     := cNumero
	    PA2->PA2_STATUS := "1"
	    PA2->PA2_OK     := oMark:Mark()
	    PA2->PA2_DATA   := dDataBase
	    PA2->PA2_FILE   := Alltrim(aArq[nX1,1])
	    PA2->PA2_LINE   := STRZERO(1,TamSx3("PA2_LINE")[1])
	    For nX3 := 1 to LEN(aFields)
	        PA2->&(aFields[nX3,1]) := aDados[nX2,nX3]
	    Next nX2  
        cMsg            := ""
        cMsg := U_ALLVERIF("PA2")
        If Alltrim(cMsg) <> ""
           PA2->PA2_ERRO   := Alltrim(cMsg)
           PA2->PA2_STATUS := "0"
        Endif    
	    PA2->(MsUnlock())                      
	    nLin++
	    If PA2->PA2_STATUS = "1"
	       lGera := .T.
	    Endif   
	Next nX2    
	FT_FUse() 
	WFMoveFiles( cPathTmp+aArq[nX1,1],cPathTmp+"Importados") 
	If lGera
       Processa({|| PA2_GERA(2)},"Gerando Títulos a Receber")
    Endif
Next nX1   
RestArea( axArea )
dbCloseAll()
cEmpAnt := cEmpBKP
cFilAnt := cFilBKP
OpenFile(cEmpAnt+cFilAnt)
OpenSM0()
DbSelectArea("PA2")
oMark:Refresh()
return 

//===============================================================
Static Function PA2_AVALIAR()    //gerando Títulos a Receber
//===============================================================
Local aSE1      := {}
Local nX1       := 0
Local nX2       := 0
Local nItem     := 0
Local cNumtit   := ""
Local axLinha   := {}
Local aItens    := {}
Local aCabec    := {}
Local aPA2      := {}
Local nXX       := 0
Local nYY       := 0
Local nZZ       := 0
Local cTexto    := 0
Local nTotal    := 0
Local cQryAux   := ""
Local cStartPah := GetSrvProfString("Startpath","")
Local cLog		:= ''
Local cMsg      := ""
Local cArqLog   := 'LOGCSV.LOG'
 
If Select("QRY_AUX")>0
   DbSelectArea("QRY_AUX")
   DbCloseArea()
Endif
cQryAux := " SELECT * FROM "+RetSqlName("PA2")
cQryAux += " WHERE D_E_L_E_T_ = ' '
cQryAux += " AND PA2_FILIAL   = '"+xFilial("PA2")+"'"     
cQryAux += " AND PA2_STATUS IN('0','1',' ')"
cQryAux += " ORDER BY PA2_ID"  
cQryAux := ChangeQuery(cQryAux)
TCQuery cQryAux New Alias "QRY_AUX"
dbselectarea("QRY_AUX")                   
Count to nCount 
QRY_AUX->(DbGoTop())
ProcRegua(nCount)
While QRY_AUX->(!EOF())  
      nItem ++   
      cMsg := ""   
      dbSelectArea("PA2")
      PA2->(dbSetOrder(1))
      PA2->(dbSeek(xFilial("PA2")+QRY_AUX->PA2_ID))
      RecLock("PA2",.F.)
      cMsg := U_ALLVERIF("PA2")
        IncProc("Avaliando ==> "+Alltrim(Str(nItem))+" de "+Alltrim(Str(nTotal))+" - "+PA2->PA2_ID+"-"+Alltrim(PA2->PA2_FANTAS))
        If Alltrim(cMsg) <> ""
           PA2->PA2_ERRO   := Alltrim(cMsg)
           PA2->PA2_STATUS := "0"
        Else
           PA2->PA2_ERRO   := " "
           PA2->PA2_STATUS := "1"          
        Endif    
	    PA2->(MsUnlock())                      
        QRY_AUX->(dbSkip())
Enddo
Return

//===============================================================
Static Function PA2_GERA(p_par)    //gerando Títulos a Receber
//===============================================================
Local aSE1      := {}
Local nX1       := 0
Local nX2       := 0
Local nItem     := 0
Local cNumtit   := ""
Local axLinha   := {}
Local aItens    := {}
Local aCabec    := {}
Local aPA2      := {}
Local nXX       := 0
Local nYY       := 0
Local nZZ       := 0
Local cTexto    := 0
Local nCount    := 0
Local cQryAux   := ""
Local cLog		:= ''
Local cxyPrefixo := ""
Local cArqLog   := 'LOGCSV.LOG'

 cMarca   := oMark:Mark()
lInverte  := oMark:IsInvert()
cPrefxTel := U_ETX_PREFIXO()
If Alltrim(cPrefxTel) = ""
   MsgStop("Prefixo (Z3) não encontrado:"+cEmpAnt+"-"+"-"+cFilAnt,"Erro")
   Return
Endif  
If p_par = 2
   dbSelectArea("PA2")
   PA2->(dbGoTop())
   While !PA2->( Eof() )
         If PA2->PA2_STATUS = "1" .AND. Alltrim(cPrefxTel) = Left(Alltrim(PA2->PA2_FILE),2) .AND. DTOS(PA2->PA2_DATA) = DTOS(dDATABASE)
            RecLock("PA2",.F.)
            PA2->PA2_OK := cMarca
            PA2->(MsUnlock())
         Endif   
         PA2->(dbSkip())
   EndDo
Endif
nItem := 0
dbSelectArea("PA2")
PA2->(dbGoTop())
While !PA2->( Eof() ) 
      nItem ++
      cTexto := ""
      If PA2->PA2_STATUS = "1" .AND. Alltrim(cPrefxTel) = Left(Alltrim(PA2->PA2_FILE),2) .AND. PA2->PA2_OK = cMarca  
         cNumtit := fLerSE1()  
         IncProc("Gerando Título ==> "+Alltrim(cEmpAnt)+"-"+Alltrim(cFilAnt)+"-"+cNumtit+"-"+Alltrim(SA1->A1_NREDUZ))
         aSE1 := {}
         dbSelectArea("SA1")
         SA1->(dbSetOrder(3))
         If !(SA1->(dbSeek(xFilial("SA1")+PA2->PA2_CNPJ)))
            cTexto += "Cliente não cadastrado"+Chr(13)+Chr(10) 
         Endif    
         If SA1->A1_MSBLQL = "1"  
            cTexto += "Cliente Inativo"+Chr(13)+Chr(10) 
          Endif  
         dbSelectArea("SB1")
         SB1->(dbSetOrder(1))
         If !(SB1->(dbSeek(xFilial("SB1")+PA2->PA2_PRODUT)))
             cTexto += "Produto não cadastrado"+Chr(13)+Chr(10) 
         Endif 
         If Alltrim(SB1->B1_XNAT) = ""
            cTexto += "Natureza Financeira não informarmada no Produto"+Chr(13)+Chr(10) 
         Endif 
         If Alltrim(cTexto) <> ""
            Reclock("PA2",.F.)
            PA2->PA2_SATUS := "0"
            PA2->PA2_OK    := " "
            PA2->PA2_ERRO := Alltrim(cTexto)
            PA2->(MsUnlock())
            PA2->(dbSkip())
            Loop
         Endif   
         aAdd(aSE1,{"E1_FILIAL"  ,xFilial("SE1")                                                                 ,Nil})
         aAdd(aSE1,{"E1_NUM"     ,cNumtit                                                                        ,Nil})
         aAdd(aSE1,{"E1_PREFIXO","IMP"                                                                          ,Nil})
         aAdd(aSE1,{"E1_PARCELA" ,"1"                                                                            ,Nil})
         aAdd(aSE1,{"E1_TIPO"    ,"IMP"                                                                          ,Nil})
         aAdd(aSE1,{"E1_NATUREZ" ,SB1->B1_XNAT                                                                   ,Nil})
         aAdd(aSE1,{"E1_CLIENTE" ,SA1->A1_COD                                                                    ,Nil})
         aAdd(aSE1,{"E1_LOJA"    ,SA1->A1_LOJA                                                                   ,Nil})
         aAdd(aSE1,{"E1_NOMCLI"  ,SA1->A1_NREDUZ                                                                 ,Nil})
         aAdd(aSE1,{"E1_EMISSAO" ,PA2->PA2_EMISSA                                                                ,Nil})
         aAdd(aSE1,{"E1_VENCTO"  ,PA2->PA2_VENCTO                                                                ,Nil})
         aAdd(aSE1,{"E1_VENCREA" ,PA2->PA2_VENCTO                                                                ,Nil})
         aAdd(aSE1,{"E1_VALOR"   ,PA2->PA2_VALOR                                                                 ,Nil})
         aAdd(aSE1,{"E1_SALDO"   ,PA2->PA2_VALOR                                                                 ,Nil})
         aAdd(aSE1,{"E1_HIST"   ,LEFT(ALLTRIM(PA2->PA2_ID)+'-'+ALLTRIM(PA2->PA2_PERIOD)+'-'+ALLTRIM(PA2->PA2_FILE),TamSx3('E1_HIST')[1]),Nil})
         aAdd(aSE1,{"E1_XPROD"   ,PA2->PA2_PRODUT                                                                ,Nil})
         aAdd(aSE1,{"E1_XREDE"   ,fLerSX5("Z1",PA2->PA2_REDE)    ,Nil})
         aAdd(aSE1,{"E1_XMASTER" ,fLerSX5("Z2",PA2->PA2_MASTER)   ,Nil})
         If cEmpAnt = "16"
            aAdd(aSE1,{"E1_XCCONT", AllTrim(SB1->B1_CTA002)                                                      ,Nil})
         Endif   
         aAdd(aSE1, {"E1_MOEDA"  ,   1,                                                                          ,Nil})
         aAdd(aSE1,{"E1_ORIGEM"  ,"FINA040"                                                                     ,Nil})
         lMsErroAuto := .F.
         MSExecAuto({|x,y| Fina040(x,y)},aSE1,3)  // 3 - Inclusao, 4 - Alteração, 5 - Exclusão
         If lMsErroAuto  
	        cStartPath  := GetSrvProfString("Startpath","")
	        cLog	    := ""	
	        cArqLog     := 'LOGCSV.LOG'
	        If File( cStartPath+cArqLog )
               Ferase(cStartPath+cArqLog)
	        Endif
	        Mostraerro(cStartPath,cArqLog)
	        If File(cStartPath+cArqLog )
               cLog := MemoRead(cStartPath+cArqLog)
               Reclock("PA2",.F.)
	           PA2->PA2_STATUS := "0"
	           PA2->PA2_OK     := "  "
	           PA2->PA2_ERRO   := cLog
	           PA2->PA2_TITULO := " "
	           PA2->(MsUnlock())
	        Endif
	     Else
	        dbSelectArea("SE1")
            SE1->(dbSetOrder(2))
            If SE1->(dbSeek(xFilial("SE1")+SA1->A1_COD +SA1->A1_LOJA +PADR("IMP",TamSx3('E1_PREFIXO')[1])+Padr(cNumtit,TamSx3('E1_NUM')[1])+PADR("1",TamSx3('E1_PARCELA')[1])+PADR("IMP",TamSx3('E1_TIPO')[1])))
	           Reclock("PA2",.F.)
	           PA2->PA2_STATUS := "2"
	           PA2->PA2_OK     := "  "
	           PA2->PA2_ERRO   := SPACE(TamSx3('PA2_ERRO')[1])
	           PA2->PA2_TITULO := cNumtit
	           PA2->(MsUnlock())
	        Endif   
         Endif   
      Endif
      PA2->(dbSkip())
Enddo
Return

//===============================================================
Static Function PA2_EXCTIT()     
//===============================================================                   
Local oAteTitulo
Local cAteTitulo := Space(TamSx3('E1_NUM')[1])
Local oButton1
Local oButton2
Local oDoTitulo
Local cDoTitulo := Space(TamSx3('E1_NUM')[1])
Local oGroup1
Local oSay1
Local oSay2
Static oDlgExcl

cDoTitulo  := PA2->PA2_TITULO
cAteTitulo := PA2->PA2_TITULO
cMsg := Space(2000)
DEFINE MSDIALOG oDlgExcl TITLE "Informe os Números dos Títulos a Receber" FROM 000, 000  TO 240, 390 COLORS 0, 16777215 PIXEL
    @ 001, 004 GROUP oGroup1 TO 028, 189 PROMPT "Informe os Números dos Títulos a Receber" OF oDlgExcl COLOR 16711680, 16777215 PIXEL
    @ 015, 009 SAY oSay1 PROMPT "De :" SIZE 020, 007 OF oDlgExcl COLORS 0, 16777215 PIXEL
    @ 012, 032 MSGET oDoTitulo VAR cDoTitulo SIZE 060, 010 OF oDlgExcl COLORS 0, 16777215 PIXEL
    @ 015, 101 SAY oSay2 PROMPT "Até :" SIZE 020, 007 OF oDlgExcl COLORS 0, 16777215 PIXEL
    @ 012, 125 MSGET oAteTitulo VAR cAteTitulo SIZE 060, 010 OF oDlgExcl COLORS 0, 16777215 PIXEL
    @ 033, 004 GET oMsg VAR cMsg OF oDlgExcl MULTILINE SIZE 185, 064 COLORS 0, 16777215 READONLY HSCROLL NOBORDER PIXEL
    @ 102, 110 BUTTON oButton1 PROMPT "Confirmar" SIZE 037, 012 OF oDlgExcl ACTION PA2_PVEXCL(cDoTitulo,cAteTitulo) PIXEL 
    @ 102, 151 BUTTON oButton2 PROMPT "Sair" SIZE 037, 012 OF oDlgExcl ACTION oDlgExcl:End() PIXEL
ACTIVATE MSDIALOG oDlgExcl CENTERED
Return

//===============================================================
Static Function PA2_PVEXCL(cDoTitulo,cAteTitulo) 
//===============================================================
Local aSE1      := {}
Local nX1       := 0
Local nX2       := 0
Local nItem     := 0
Local cNumtit   := ""
Local axLinha   := {}
Local aItens    := {}
Local aCabec    := {}
Local aPA2      := {}
Local nXX       := 0
Local nYY       := 0
Local nZZ       := 0
Local cTexto    := 0
Local nCount    := 0
Local cQryAux   := ""
Local cStartPah := GetSrvProfString("Startpath","")
Local cLog		:= ''
Local cArqLog   := 'LOGCSV.LOG'

IF Select("QRY_AUX") > 0 
   DbSelectArea("QRY_AUX")
   DbCloseArea() 
Endif
cQryAux := " SELECT * FROM "+RetSqlName("PA2")+ XENTERX
cQryAux += " WHERE D_E_L_E_T_ = ' '"+ XENTERX
cQryAux += " AND PA2_FILIAL   = '"+xFilial("PA2")+"'"+ XENTERX    
cQryAux += " AND PA2_STATUS   = '2'"+ XENTERX
cQryAux += " AND PA2_TITULO BETWEEN '"+cDoTitulo+"' AND '"+cAteTitulo+"'"+ XENTERX
cQryAux += " ORDER BY PA2_ID"  
cQryAux := ChangeQuery(cQryAux) 
TCQuery cQryAux New Alias "QRY_AUX" 	
dbselectarea("QRY_AUX")                   
Count to nCount 
QRY_AUX->(DbGoTop())
ProcRegua(nCount)
While QRY_AUX->(!EOF())   
      cPrefxTel := U_ETX_PREFIXO() 
      If Alltrim(cPrefxTel) <> Left(Alltrim(QRY_AUX->PA2_FILE),2)
         QRY_AUX->(dbSkip())
         Loop
      Endif   
      nItem ++
      dbSelectArea("PA2")
      PA2->(dbSetOrder(1))
      If PA2->(dbSeek(xFilial("PA2")+QRY_AUX->PA2_ID))
         IncProc("Processando ==> "+Alltrim(Str(nItem))+" de "+Alltrim(Str(nCount))+" - "+cNumtit+"-"+Alltrim(SA1->A1_NREDUZ))
         aSE1 := {}
         dbSelectArea("SA1")
         SA1->(dbSetOrder(3))
         If !(SA1->(dbSeek(xFilial("SA1")+QRY_AUX->PA2_CNPJ)))
            MsgStop ("Cliente não cadastrado: "+QRY_AUX->PA2_CNPJ,"Erro")
            Return
         Endif    
         If SA1->A1_MSBLQL = "1"  
            MsgStop ("Cliente Inativo: "+SA1->A1_COD+"-"+SA1->A1_LOJA,"Erro")
            Return
          Endif  
         dbSelectArea("SB1")
         SA1->(dbSetOrder(1))
         If !(SB1->(dbSeek(xFilial("SB1")+QRY_AUX->PA2_PRODUT)))
             MsgStop ("Produto não cadastrado: "+QRY_AUX->PA2_PRODUT,"Erro")
             Return
         Endif 
         If Alltrim(SB1->B1_XNAT) = ""
            MsgStop ("Natureza Financeira não informarmada no Produto: "+QRY_AUX->PA2_PRODUT,"Erro")
             Return
         Endif 
         dbSelectArea("SE1")
         SE1->(dbSetOrder(2))
         If SE1->(dbSeek(xFilial("SE1")+SA1->A1_COD +SA1->A1_LOJA +PADR("IMP",TamSx3('E1_PREFIXO')[1])+PA2->PA2_TITULO+PADR("1",TamSx3('E1_PARCELA')[1])+PADR("IMP",TamSx3('E1_TIPO')[1])))
            aAdd(aSE1,{"E1_FILIAL" ,xFilial("SE1")                     ,Nil})
            aAdd(aSE1,{"E1_NUM"    ,PA2->PA2_TITULO                    ,Nil})
            aAdd(aSE1,{"E1_PREFIXO",PADR("IMP",TamSx3('E1_PREFIXO')[1]),Nil}) 
            aAdd(aSE1,{"E1_PARCELA",PADR("1",TamSx3('E1_PARCELA')[1])  ,Nil}) 
            aAdd(aSE1,{"E1_TIPO"   ,PADR("IMP",TamSx3('E1_TIPO')[1])   ,Nil})
            aAdd(aSE1,{"E1_NATUREZ",SB1->B1_XNAT                       ,Nil})
            aAdd(aSE1,{"E1_CLIENTE",SA1->A1_COD                        ,Nil})
            aAdd(aSE1,{"E1_LOJA"   ,SA1->A1_LOJA                       ,Nil})
            aAdd(aSE1,{"E1_NOMCLI" ,SA1->A1_NREDUZ                     ,Nil})
            aAdd(aSE1,{"E1_EMISSAO",PA2->PA2_EMISSA                    ,Nil})
            aAdd(aSE1,{"E1_VENCTO" ,PA2->PA2_VENCTO                    ,Nil})
            aAdd(aSE1,{"E1_VENCREA",PA2->PA2_VENCTO                    ,Nil})
            aAdd(aSE1,{"E1_VALOR"  ,PA2->PA2_VALOR                     ,Nil})
            aAdd(aSE1,{"E1_HIST"   ,LEFT(ALLTRIM(PA2->PA2_ID)+'-'+ALLTRIM(PA2->PA2_PERIOD)+'-'+ALLTRIM(PA2->PA2_FILE),TamSx3('E1_HIST')[1]),Nil})
            aAdd(aSE1,{"E1_XPROD"  ,PA2->PA2_PRODUT                    ,Nil})
            aAdd(aSE1,{"E1_XREDE"  ,PA2->PA2_REDE                      ,Nil})
            aAdd(aSE1,{"E1_XMASTER",PA2->PA2_MASTER                    ,Nil}) 
            aAdd(aSE1, {"E1_MOEDA" ,   1,                              ,Nil})
            aAdd(aSE1,{"E1_ORIGEM" ,"FINA040"                          ,Nil})
	        lMsErroAuto := .F.
            MsExecAuto({|x,y| FINA040(x,y)},aSE1, 5)  // 3 - Inclusao, 4 - Alteração, 5 - Exclusão
            If lMsErroAuto //ALLIMPFIN
               //Mostraerro()
	           cStartPath  := GetSrvProfString("Startpath","")
	           cLog	    := ""	
	           cArqLog     := 'LOGCSV.LOG'
	           If File( cStartPath+cArqLog )
                  Ferase( cStartPath+cArqLog )
	           Endif
	           Mostraerro(cStartPath,cArqLog)
	           If File(cStartPath+cArqLog )
                  cLog := MemoRead(cStartPath+cArqLog)
                  cMsg += cLog+Chr(13)+Chr(10)
                  Reclock("PA2",.F.)
	              PA2->PA2_ERRO := cLog
	              PA2->(MsUnlock())
	           Endif
	        Else
	           Reclock("PA2",.F.)
	           PA2->PA2_STATUS := "1"
	           PA2->PA2_ERRO   := ""
	           PA2->PA2_TITULO := ""
	           PA2->(MsUnlock())
	           cMsg += "Excluído Título: "+PA2->PA2_TITULO+" para o ID: "+PA2->PA2_ID+" gerado com sucesso"+Chr(13)+Chr(10)
            Endif
         Endif   
      Endif   
      oMsg:Refresh()
      oDlgExcl:Refresh()
      QRY_AUX->(dbSkip())
Enddo
oMsg:Refresh()
oDlgExcl:Refresh()
Return

//===============================================================
Static Function PA2_TELOTE(cDoTitulo,cAteTitulo)     
//===============================================================                   
Local oAtePedido
Local cAteTitulo := Space(TamSx3('E1_NUM')[1])
Local oButton1
Local oButton2
Local oDoPedido
Local cDoTitulo  := Space(TamSx3('E1_NUM')[1])
Local oGroup1
Local oSay1
Local oSay2
Static oDlgExcl

cDoTitulo  := PA2->PA2_ID
cAteTitulo := PA2->PA2_ID
cMsg := Space(2000)
DEFINE MSDIALOG oDlgExcl TITLE "Informe o intervalo dos Id´s" FROM 000, 000  TO 240, 390 COLORS 0, 16777215 PIXEL
    @ 001, 004 GROUP oGroup1 TO 028, 189 PROMPT "Informe os Números dos ID´s" OF oDlgExcl COLOR 16711680, 16777215 PIXEL
    @ 015, 009 SAY oSay1 PROMPT "De :" SIZE 020, 007 OF oDlgExcl COLORS 0, 16777215 PIXEL
    @ 012, 032 MSGET oDoPedido VAR cDoTitulo SIZE 060, 010 OF oDlgExcl COLORS 0, 16777215 PIXEL
    @ 015, 101 SAY oSay2 PROMPT "Até :" SIZE 020, 007 OF oDlgExcl COLORS 0, 16777215 PIXEL
    @ 012, 125 MSGET oAtePedido VAR cAteTitulo SIZE 060, 010 OF oDlgExcl COLORS 0, 16777215 PIXEL
    @ 033, 004 GET oMsg VAR cMsg OF oDlgExcl MULTILINE SIZE 185, 064 COLORS 0, 16777215 HSCROLL NOBORDER PIXEL
    //@ 102, 110 BUTTON oButton1 PROMPT "Confirmar" SIZE 037, 012 OF oDlgExcl ACTION Processa({|| PA2_PVEXCL(cDoTitulo,cAteTitulo)},"Imprimindo Relatório") PIXEL 
    @ 102, 110 BUTTON oButton1 PROMPT "Confirmar" SIZE 037, 012 OF oDlgExcl ACTION PA2_LOTE(cDoTitulo,cAteTitulo) PIXEL 
    @ 102, 151 BUTTON oButton2 PROMPT "Sair" SIZE 037, 012 OF oDlgExcl ACTION oDlgExcl:End() PIXEL
ACTIVATE MSDIALOG oDlgExcl CENTERED
Return

//===============================================================
Static Function PA2_LOTE(cDoTitulo,cAteTitulo)
//===============================================================
Local nXX     := 0
Local nCount  := 0
Local cQryAux := ""
Local aCabec  := {}
Local aItens  := {}
Local axLinha  := {}
Local aPA2    := {}
Local nTotal  := 0
Local nAtual  := 0

cMsg += "=====> Processamento iniciado "+ XENTERX
cMsg += ""+ XENTERX
If Select("QRY_AUX")>0
   DbSelectArea("QRY_AUX")
   DbCloseArea()
Endif
oMsg:Refresh()
oDlgExcl:Refresh()
cQryAux   := ""
cQryAux += " SELECT * FROM "+RetSqlName("PA2")+ XENTERX
cQryAux += " WHERE D_E_L_E_T_ = ''"	+ XENTERX
cQryAux += " AND PA2_FILIAL = '"+xFilial("PA2")+"'"+XENTERX
cQryAux += " AND PA2_STATUS IN ('0','1')"+ XENTERX
cQryAux += " AND PA2_ID BETWEEN '"+cDoTitulo+"' AND '"+cAteTitulo+"'"+XENTERX
cQryAux += " AND PA2_TITULO = ' '"+ XENTERX
cQryAux += " ORDER BY PA2_ID"+ XENTERX
cQryAux := ChangeQuery(cQryAux)
TCQuery cQryAux New Alias "QRY_AUX"
Count to nTotal
ProcRegua(nTotal)
QRY_AUX->(DbGoTop())
While !QRY_AUX->(Eof())
      nAtual++
      dbSelectArea("PA2")
      PA2->(dbSetOrder(1))
      PA2->(dbSeek(xFilial("PA2")+QRY_AUX->PA2_ID))
      RecLock("PA2",.F.)
      PA2->(dbDelete())
      PA2->(MsUnlock())
      cMsg += "----> Registro Excluído com sucesso : "+PA2->PA2_ID+" - "+PA2->PA2_FANTAS+ XENTERX
      oMsg:Refresh()
      oDlgExcl:Refresh()
      QRY_AUX->(DbSkip())      
EndDo
cMsg += ""+ XENTERX
cMsg += "=====> Processamento finalizado "+ XENTERX
oMsg:Refresh()
oDlgExcl:Refresh()
Return

//====================================================================================
Static Function fLerSE1()
//====================================================================================
Local cTitAux := ""

If Select("RS_SE1") > 0 
   DbSelectArea("RS_SE1")
   DbCloseArea() 
Endif
cQryAux := " SELECT MAX(E1_NUM) E1_NUM "
cQryAux += " FROM " + RetSqlName('SE1')
cQryAux += " WHERE E1_PREFIXO = 'IMP' "
cQryAux += " AND E1_TIPO = 'IMP' "
cQryAux := ChangeQuery( cQryAux )
TCQuery cQryAux New Alias "RS_SE1" 		
dbselectarea("RS_SE1")           
RS_SE1->(DbGoTop())       
While (!RS_SE1->(EOF()))  
      cTitAux := STRZERO(VAL(RS_SE1->E1_NUM)+1,TamSx3('PA2_TITULO')[1])
      RS_SE1->(DbSkip())       
EndDo 
If Alltrim(cTitAux) = ""
   cTitAux := STRZERO(1,TamSx3('E1_TITULO')[1])
Endif
Return (cTitAux)

//====================================================================================
Static Function fLerSA1(p_CNPJ)
//====================================================================================
Local cQryAux := ""
Local cCnpj  := p_CNPJ
Local nRecno := 0

If Select("TMPSA1")>0
   DbSelectArea("TMPSA1")
   DbCloseArea()
Endif
cQryAux := " SELECT * FROM " + RetSqlName('SA1')
cQryAux += " WHERE D_E_L_E_T_ = ' '  "
cQryAux += " AND A1_FILIAL = '"+xFilial("SA1")+"'"
cQryAux += " AND A1_CGC LIKE '%"+ALLTRIM(cCnpj)+"%'"
cQryAux += " AND ORDER BY A1_MSBLQL"
//dbUseArea(.T., "TOPCONN", TCGENQRY(,,cQryAux), "TMPSA1", .F., .T.)
TCQuery cQryAux New Alias "TMPSA1"	
While !TMPSA1->( Eof() )
    nRecno := TMPSA1->A1_RECNO
	cCnpj := TMPSA1->A1_CGC
	TMPSA1->(dbSkip())
Enddo
Return ({cCnpj,nRecno})

//====================================================================================
Static Function fLerSX5(cTabela,cDescri)
//====================================================================================
Local cQryAux := ""
Local cxxRet := ""

If Select("TMPSX5")>0
   DbSelectArea("TMPSX5")
   DbCloseArea()
Endif
cQryAux := " SELECT X5_CHAVE "
cQryAux += " FROM " + RetSqlName('SX5')
cQryAux += " WHERE D_E_L_E_T_ = ' '  "
cQryAux += " AND X5_TABELA  = '"+cTabela+"'"
cQryAux += " AND X5_DESCRI LIKE '%"+AllTrim(cDescri)+"%'"
cQryAux := ChangeQuery(cQryAux) 
TCQuery cQryAux New Alias "TMPSX5" 
If !TMPSX5->(Eof()) .and. !TMPSX5->(Bof())
	cxxRet := TMPSX5->X5_CHAVE
Endif
Return (cxxRet)

//------------------------------------------------------------------------------------------
Static Function PA2_DOCUMENTO()
//------------------------------------------------------------------------------------------

cStartPath   := If(cEmpAnt="99",'C:\TOTVS12\protheus_data\Cisao-Log\','E:\TOTVS_12_EC_HML_CISAO\protheus_data\Cisao-Log\')
cArqLog      := "PC-Cisao empresas-Importação deTítulos a Receber.DOCX"
cChave       := Alltrim(cStartPath)+Alltrim(cArqLog)
ShellExecute("open",Alltrim(cArqLog),"",Alltrim(cStartPath),3)

Return(.T.)
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

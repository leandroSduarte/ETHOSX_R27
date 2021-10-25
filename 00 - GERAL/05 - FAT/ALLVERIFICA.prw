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
User Function ALLVERIF(cXAlias)
//=============================================================
Local cMsgErro := ""
Local aCNPJ    := {}
Local nRecno   := 0
Local cPrefixo := If(Left(cXAlias,1)="S",Substr(cXAlias,2,2),cXAlias)

cPrefixo := If(Left(cXAlias,1)="S",Substr(cXAlias,2,2),cXAlias)
If (cXAlias)->&(cPrefixo+"_CNPJ") = ""
   cMsgErro += "Cnpj inválido" +Chr(13)+Chr(10)
Else   
   If "E|+" $ (cXAlias)->&(cPrefixo+"_CNPJ")
      cMsgErro += "Cnpj com caracter inválido"+Chr(13)+Chr(10)
   Endif   
   If VAL((cXAlias)->&(cPrefixo+"_CNPJ")) = 0
      cMsgErro += "Cnpj inválido"+Chr(13)+Chr(10)
   Endif  
Endif 
If (cXAlias)->&(cPrefixo+"_CNPJ") <> ""
   aCNPJ := fLerSA1((cXAlias)->&(cPrefixo+"_CNPJ"))
   (cXAlias)->&(cPrefixo+"_CNPJ") := aCNPJ[1]
   nRecno        := aCNPJ[2]
   If nRecno = 0
      cMsgErro += "Cliente"+Chr(13)+Chr(10)
   Else  
      SA1->(dbGoTo(nRecno))  
      If SA1->A1_MSBLQL = "1"  
         cMsgErro += "Cliente Inativo"+Chr(13)+Chr(10)
      Endif   
   Endif      
Endif
dbSelectArea("SB1")
SB1->(dbSetOrder(1))
If !(SB1->(dbSeek(xFilial("SB1")+(cXAlias)->&(cPrefixo+"_PRODUT"))))  
   cMsgErro += "Produto"+Chr(13)+Chr(10)
Else   
   If Alltrim(cXAlias) = "PPX"
      PPX->PPX_CONDPG  := SB1->B1_XCOND
      PPX->PPX_CODNAT  := SB1->B1_XNAT
   Endif   
Endif 
If !(fLerSX5("Z2",(cXAlias)->&(cPrefixo+"_MASTER")))
   cMsgErro += "Master"+Chr(13)+Chr(10)
Endif
If !(fLerSX5("Z1",(cXAlias)->&(cPrefixo+"_REDE")))
   cMsgErro += "Rede"+Chr(13)+Chr(10)
Endif   
If Round((cXAlias)->&(cPrefixo+"_VALOR"),2) <= Round((cXAlias)->&(cPrefixo+"_DESC"),2)
   cMsgErro += "Descontos"+Chr(13)+Chr(10)
Endif   
If Round((cXAlias)->&(cPrefixo+"_VALOR"),2) = 0
   cMsgErro += "Valor"+Chr(13)+Chr(10)
Endif  
If Alltrim((cXAlias)->&(cPrefixo+"_FANTAS")) = ""
   cMsgErro += "Fantasia"+Chr(13)+Chr(10)
Endif  
If Alltrim((cXAlias)->&(cPrefixo+"_RAZAO")) = ""
   cMsgErro += "Razão Social"+Chr(13)+Chr(10)
Endif  
If Alltrim((cXAlias)->&(cPrefixo+"_PERIOD")) = ""
   cMsgErro += "Período"+Chr(13)+Chr(10)
Endif
If Alltrim(DTOS((cXAlias)->&(cPrefixo+"_EMISSA"))) = ""
   cMsgErro += "Dt Emissão"+Chr(13)+Chr(10)
Endif
If Alltrim(DTOS((cXAlias)->&(cPrefixo+"_VENCTO"))) = ""
   cMsgErro += "Dt Vencimento"+Chr(13)+Chr(10)
Endif
If Alltrim(DTOS((cXAlias)->&(cPrefixo+"_VENCTO"))) < Alltrim(DTOS((cXAlias)->&(cPrefixo+"_EMISSA"))) 
   cMsgErro += "Dt Vencimento inválida"+Chr(13)+Chr(10)
Endif
If Alltrim(cXAlias) = "PPX"
   dbSelectArea("SE4")
   SE4->(dbSetOrder(1))
   If !(SE4->(dbSeek(xFilial("SE4")+SB1->B1_XCOND)))   
      cMsgErro += "Condição de Pagamento"+Chr(13)+Chr(10)
   Endif  
   dbSelectArea("SED")
   SED->(dbSetOrder(1))
   If !(SED->(dbSeek(xFilial("SED")+SB1->B1_XNAT)))   
      cMsgErro += "Natureza Financeira"+Chr(13)+Chr(10)
   Endif
Endif 
Return (cMsgErro)

//====================================================================================
Static Function fLerSA1(p_CNPJ)
//====================================================================================
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

//====================================================================================
Static Function fLerSX5(cTabela,cDescri)
//====================================================================================
Local lAchou := .F.

If Select("TMPSX5")>0
   DbSelectArea("TMPSX5")
   DbCloseArea()
Endif
cQryAux := " SELECT * FROM " + RetSqlName('SX5')
cQryAux += " WHERE D_E_L_E_T_ = ' '  "
cQryAux += " AND X5_TABELA  = '"+cTabela+"'"
If cTabela = "Z2"
   cQryAux += " AND UPPER(X5_DESCRI) LIKE '%"+UPPER(AllTrim(cDescri))+"%'"
Else
   cQryAux += " AND UPPER(X5_DESCRI) = '"+UPPER(AllTrim(cDescri))+"'"   
Endif   
TCQuery cQryAux New Alias "TMPSX5"
If !TMPSX5->(Eof()) .and. !TMPSX5->(Bof())
	lAchou := .T.
Endif
Return (lAchou)

//====================================================================================
User Function ETX_EMPRE(p_Prefix)
//====================================================================================
Local cQuery   := ""
Local cTitulo  := ""
Local cNomex   := ""
Local cPrefix  := ""
Local cEmpresx  := ""
Local cFilialx  := ""

If Select("TMPEMP")>0
   DbSelectArea("TMPEMP")
   DbCloseArea()
Endif
cQuery := " SELECT * FROM " + RetSqlName('SX5')
cQuery += " WHERE D_E_L_E_T_ = ' '  "
cQuery += " AND X5_TABELA  = 'Z3'"
cQuery += " AND UPPER(X5_CHAVE) = '"+UPPER(ALLTRIM(p_Prefix))+"'"
cQuery += " AND X5_DESCSPA  <> 'Z3'"
TCQuery cQuery New Alias "TMPEMP"
If !TMPEMP->(Eof())
	cPrefix  := Alltrim(TMPEMP->X5_CHAVE)
	cNomex   := Alltrim(TMPEMP->X5_DESCRI)
	cEmpresx := Alltrim(TMPEMP->X5_DESCSPA)
	cFilialx := Alltrim(TMPEMP->X5_DESCENG)
	cTitulo  := Alltrim(TMPEMP->X5_DESCRI)+"-"+Alltrim(TMPEMP->X5_CHAVE)+"-"+Alltrim(TMPEMP->X5_DESCRI)
	TMPEMP->(dbSkip()) 
Endif
Return ({cPrefix,cEmpresx,cFilialx,cNomex,cTitulo})

//============================================================
User Function ETX_PREFIXO()
//============================================================
Local cPrefxTel := ""

If Select("TMPSX5")>0
   DbSelectArea("TMPSX5")
   DbCloseArea()
Endif
cQryAux := " SELECT * FROM " + RetSqlName('SX5')
cQryAux += " WHERE D_E_L_E_T_ = ' '  "
cQryAux += " AND X5_TABELA  = 'Z3'"
cQryAux += " AND X5_DESCSPA  = '"+cEmpAnt+"'"
cQryAux += " AND X5_DESCENG  = '"+cFilAnt+"'"
TCQuery cQryAux New Alias "TMPSX5"
If !TMPSX5->(Eof()) .and. !TMPSX5->(Bof())
	cPrefxTel := TMPSX5->X5_CHAVE
Else
   cQryAux := "Não foi encontrado (Z3)-Cadastro de Prefixo de Arquivos para:"+Chr(13)+Chr(10)
   cQryAux += "Empresa :"+cEmpAnt+Chr(13)+Chr(10)
   cQryAux += "Filial  :"+cFilAnt+Chr(13)+Chr(10)
   MsgStop(cQryAux,"Atenção")	
Endif
If Alltrim(cPrefxTel) <> ""
   cLog := "UPDATE "+RetSqlName("PPX")+" SET PPX_PREFIX = '"+cPrefxTel+"' WHERE PPX_PREFIX = ' '"
   TCSQLExec(cLog)
   cLog := "UPDATE "+RetSqlName("PPX")+" SET PPX_PREFIX = 'PP' WHERE PPX_PREFIX IN ('PX','PZ')"
   TCSQLExec(cLog)
Endif  
Return(cPrefxTel)
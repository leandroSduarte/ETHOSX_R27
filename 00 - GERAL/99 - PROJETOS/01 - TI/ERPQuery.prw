#include 'protheus.ch'
#include 'parmtype.ch'
#include 'TbiConn.ch'
#include 'TopConn.ch'
#include 'TryException.ch'
#define CRLF (chr(13)+chr(10))

/*/{Protheus.doc} ERPQuery
//TODO Interface para Gerar Query´s de consulta com extração para Microsoft Excel.
@author Totvs - Juliano Souza
@since 23/0/2017
@version 1.0
@return nil, ${return_description}

@type function
/*/
User Function ERPQuery()

	Local oERPQry
	Local cERPQry   := Space(1)
	Local cMsg		:= ""
	Local nRegMax   := 1000
	Local lMaster   := .F.
	Local cGruposNm := ""
	Local aGruposNm := UsrRetGrp(cUserName)
	Local AliasExcl
	Private oERPResu
	Private cERPResu  := "Exemplo de Referência em Tabelas Internas = Select * From SM4" + Alltrim(cEmpAnt) + "0
	Private oDlg
	Private oGetDB
	PRivate cQry    	:= getNextAlias()
	Private cUsrRep	:= SuperGetMV("HB_QUERYUS",,"000000") // Usuários que possuem acesso a rotina Reprocessar

	If __cUserID $ cUsrRep // Validação de usuários que possuem acesso ao reprocessamento

		// ETHOSX - MOA - 10/03/2020 - Solicitado pelo Gerente Eduardo Matias em 09/03/2020.
		lMaster := .T.
		/*
		If Len(aGruposNm) > 0
		cGruposNm := aGruposNm[1]
		EndIf

		If "000000"$cGruposNm .Or. __cuserid == "000000" 
		lMaster := .T.
		EndIf
		*/

		DEFINE MSDIALOG oDlg TITLE "Totvs ERPQuery" FROM 000,000 TO 560,900 PIXEL
		oERPQry := tMultiget():New(002,002,{|u|if(Pcount()>0,cERPQry:=u,cERPQry)},oDlg,448,090,,,,,,.T.,,,,,,!lMaster,,,,,.T.,"Expressão Query",1,,CLR_GREEN)
		@ 104,006 BUTTON "Executar (F5)"  SIZE  140,16 PIXEL OF oDlg ACTION ExecQry(@cERPQry, @cERPResu, @lMaster, oDlg:Refresh() )
		@ 104,156 BUTTON "Excel (F8)"  SIZE  140,16 PIXEL OF oDlg ACTION MsAguarde({||ExecXcel(@cERPQry,cQry)},"Aguarde","Executando Query e Excel da consulta...",.F.)
		@ 104,306 BUTTON "Sair"  SIZE  140,16 PIXEL OF oDlg ACTION oDlg:End()
		oERPResu := tMultiget():New(122,002,{|u|if(Pcount()>0,cERPResu:=u,cERPResu)},oDlg,448,27,,,,,,.T.,,,{||.F.},,,,,,,,.T.,"Log. (Empresa Conectada = " + Alltrim(cEmpAnt) + "0)",1,,CLR_RED)
		SetKey(VK_F5, {|| ExecQry(@cERPQry, @cERPResu, @lMaster, oDlg:Refresh() )})
		SetKey(VK_F8, {|| MsAguarde({||ExecXcel(@cERPQry,cQry)},"Aguarde","Executando Query e Excel da consulta...",.F.)})
		SetKey(VK_ESCAPE, {|| oDlg:End() })
		ACTIVATE MSDIALOG oDlg CENTER

	Else
		cMsg := "Você não possui acesso a execução da Rotina ERPQuery!"
		cMsg += "<br>"
		cMsg += "<br>"
		cMsg += "Caso necessário entre em contato com o Administrador do Sistema."
		cMsg += "<br>"
		cMsg += "<br>"
		cMsg += "Informe o usuário  [ <b>" + Trim(__cUserID) + " - " + Trim(UsrRetName(__cUserID))
		cMsg += " </b> ]  para agilizar o processo!"		

		MsgInfo(cMsg,"ERPQuery")
	EndIf
Return


Static Function ExecXcel(_cQuery, _aAlias)

	Local nCnt 	  := 0
	Local aStruQry  := {}
	Local cArquivo  := "ERPQuery.XLS"
	Local oExcelApp := Nil
	Local cPath     := "C:\ERPQuery\"
	Local nTotal    := 0
	Local aLinExcel := {}
	Local oExcel
	Local oExcelApp

	If ("INSERT" $ Alltrim(UPPER(_cQuery))) .OR. ("UPDATE" $ Alltrim(UPPER(_cQuery))) .OR. ("DELETE" $ Alltrim(UPPER(_cQuery))) .OR. ("CREATE" $ Alltrim(UPPER(_cQuery))) 
		Alert("Não é possível gerar planilha Excel quando a query contém os comandos: INSERT, UPDATE, DELETE ou CREATE")
		Return
	Endif

	If !Empty(_cQuery)
		TryException
		If Select(cQry) > 0
			(cQry)->(DbCloseArea())
		Endif
		DbUseArea(.T.,"TOPCONN",TcGenQry(,,_cQuery),cQry,.T.,.T.)
		(cQry)->(DbEval({|| nCnt++}))
		(cQry)->(DbGoTop())
		aStruQry  := (cQry)->(dbStruct())
		CatchException using oException
		Alert("Houve um erro na execução da Query, por favor verifique!")
		Return
		EndException
	Endif
	If nCnt <= 0
		Alert("Não ha Dados na Tabela para Gerar o EXCEL.")
		Return
	Endif
	aColunas := {}
	aLocais := {}
	oBrush1 := TBrush():New(, RGB(193,205,205))
	// Verifica se o Excel está instalado na máquina
	
	If !ApOleClient("MSExcel")
		MsgAlert("Microsoft Excel não instalado!")
		Return
	EndIf
	
	oExcel  := FWMSExcel():New()
	cAba    := "Totvs - ERPQuery"
	cTabela := "ERPQuery"
	// Criação de nova aba 
	oExcel:AddworkSheet(cAba)
	// Criação de tabela
	oExcel:AddTable (cAba,cTabela)
	// Criação de colunas 
	For nCnt := 1 To Len(aStruQry)
		oExcel:AddColumn(cAba,cTabela,aStruQry[nCnt,1],IIF(aStruQry[nCnt,2]=="N",3,1),IIF(aStruQry[nCnt,2]=="N",1,1),.F.)
	Next nCnt
	While !(cQry)->(Eof())
		// Criação de Linhas
		aLinExcel := {}
		For nCnt := 1 To Len(aStruQry)
			aAdd(aLinExcel,(cQry)->&(aStruQry[nCnt,1]))
		Next nCnt
		oExcel:AddRow(cAba,cTabela, aLinExcel)
		(cQry)->(dbSkip())
	EndDo
	If !Empty(oExcel:aWorkSheet)
		oExcel:Activate()
		oExcel:GetXMLFile(cArquivo)
		CpyS2T("\SYSTEM\"+cArquivo, cPath)
		oExcelApp := MsExcel():New()
		oExcelApp:WorkBooks:Open(cPath+cArquivo) // Abre a planilha
		oExcelApp:SetVisible(.T.)
	EndIf
Return


Static Function ExecQry(cERPQry, cERPResu, lMaster)
	Local nCnt    	:= 0
	Local cFileOpen := ""
	Local cTitulo1  := "Selecione a Query"
	Local cExtens   := "Arquivo TXT | *.txt"
	Local cLocPatc  := 'SERVIDOR\query_ERP'
	Local cBuffer	:= ""
	Local aLinha    := {}
	Local nI        := 0
	Local cAfetadas := ""
	Local lAfetadas := .T.
	Local nCont		:= 0
	Local i
	aLinha := StrToKarr(cERPQry,CHR(13)+CHR(10))
	If !Empty(cERPQry)
		If Select(cQry) > 0
			(cQry)->(DbCloseArea())
		EndIf
		If ("DELETE " $ Upper(cERPQry) .OR. "UPDATE " $ Upper(cERPQry) .OR. "INSERT INTO " $ Upper(cERPQry) .OR. "CREATE TABLE " $ Upper(cERPQry))
			If MsgYesNo("Cofirma Execução da Query?" + CRLF + CRLF + cERPQry)
				 // For i:=1 To 80
				 // nCont++
					If TcSqlExec(cERPQry) < 0
						Alert("Erro na Execução da Query de Atualização!")
						Alert(TcSqlError())
					Else
						ApMsgInfo("Atualização executada com Sucesso!")
					Endif
				// Next i
			Endif
				// Alert("Contador: "+str(nCont))
			Return
		Endif
		TryException
		//cERPQry := ChangeQuery(cERPQry)
		DbUseArea(.T.,"TOPCONN",TcGenQry(,,cERPQry),cQry,.T.,.T.)
		//(cQry)->(DbEval({|| nCnt++}))   
		(cQry)->(DbGoTop())
		MsAguarde({||GeraArq(cQry)},"Aguarde","Executando Query...",.F.)
		If Select(cQry) > 0
			(cQry)->(DbCloseArea())
		EndIf
		CatchException using oException
		Alert("Houve um erro na execução da Query, por favor verifique!")
		EndException
	Else
		MsgInfo("Comando Query não informado")
	EndIf
Return cQry


Static Function GeraArq(cQry)
	Local nRet    := 0
	Local aHeader := {}
	Local nI      := 0
	Local cArquivo 	:= CriaTrab(,.F.)
	Local cArqDBF 	:= CriaTrab(,.F.)
	Local cPath		:= AllTrim(GetTempPath())
	Local oExcelApp
	Local nHandle
	Local cCrLf 	:= Chr(13) + Chr(10)
	Local nX
	Local cDirDocs  := "C:\ERPQuery"
	Local nCntAft   := 0
	Local aStruQry	:= {}
	Local aCmpQry   := {}
	Local aResQry   := {}
	Local aResult   := {}
	Local choraIni  := Time()
	Local choraFim  := Time()
	Local cRetSql   := ""
	Local nQtdSql   := 0
	If !lIsDir(cDirDocs)
		nRet := MakeDir( cDirDocs, Nil, .F. )
		if nRet != 0
			Alert( "Não foi possível criar o diretório "+cDirDocs+", crie manualmente. Erro: " + cValToChar( FError() ) )
		endif
	Endif
	DbSelectArea(cQry)
	aStruQry := (cQry)->(dbStruct())
	For nI := 1 To Len(aStruQry)
		If aStruQry[ni,2] <> "C"
			aAdd(aCmpQry,{aStruQry[ni,1], aStruQry[ni,1], "@E", aStruQry[ni,3], aStruQry[ni,4]})
		Else
			aAdd(aCmpQry,{aStruQry[ni,1], aStruQry[ni,1], "@!", aStruQry[ni,3], aStruQry[ni,4]})
		Endif
	Next nI
	(cQry)->(DbGoTop())
	While (cQry)->(!Eof())
		nQtdSql++
		aResQry := {}
		For nI := 1 To Len(aCmpQry)
			aAdd(aResQry,(cQry)->&(aCmpQry[nI,1]))
		Next nI
		aAdd(aResQry,.F.)
		aAdd(aResult,aResQry)
		(cQry)->(DbSkip())
	EndDo
	oGetDB := MsNewGetDados():New(161,002,278,451,,,,,{''},,,,,,oDlg,aCmpQry,aResult)
	choraFim  := Time()
	cRetSql += "Exemplo de Referência em Tabelas Internas = Select * From SM4" + Alltrim(cEmpAnt) + "0" + CRLF
	cRetSql += "Tempo Execução = " + ElapTime(choraIni,choraFim) + CRLF
	cRetSql += "Qtd. Registros = " + Alltrim(Str(nQtdSql))+ CRLF
	cERPResu := cRetSql
	oERPResu:Refresh()
	oDlg:Refresh()
Return
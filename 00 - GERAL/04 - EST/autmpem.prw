#include "protheus.ch"
#include "parmtype.ch"
#include "FWBROWSE.CH"
#include "PARMTYPE.CH"
#include "FWMVCDEF.CH"
#include "FWMVCDEF.CH"

User Function AutMpEm()

	Private aObjects	:=	{}
	Private aCoors  	:=	MsAdvSize()
	Private aPosObj
	Private aCab01		:=	{}
	Private aCab02		:=	{}
	Private aItens01	:=	{}
	Private aItens02	:=	{}
	Private oLbOK		:=	LoadBitmap( GetResources(), "LBOK" )
	Private oLbNo		:=	LoadBitmap( GetResources(), "LBNO" )
	Private nAtGrid3	:=	1
	Private aProcFil	:=	{}
	Private lMark		:= .T.
	Private oFont1		:=	TFont():New("Times New Roman",,-12, .T. )
	Private oFont2		:=	TFont():New("Times New Roman",,-10, .T. )

	AAdd(aObjects,{100,100, .T. , .T. })
	aInfo:={aCoors[1],aCoors[2],aCoors[3],aCoors[4],5,5}
	aPosObj:= MsObjSize(aInfo,aObjects, .T. , .F. )

	aProcFil:= ProcFil()

	If Len(aProcFil)=0
		Iif(FindFunction("APMsgStop"), APMsgStop("Nenhuma filial foi selecionada", ".:ERRO:."), MsgStop("Nenhuma filial foi selecionada", ".:ERRO:."))
		Return()
	Else
		oProcess := MsNewProcess():New({|lEnd| LoadData(@oProcess, @lEnd) },"Analisando movimentação","Acessando banco de dados", .T. )
		oProcess:Activate()
	EndIf

	oDlg2 := MsDialog():New( aCoors[7],aCoors[7], aCoors[6], aCoors[5], "Gestão MP/ME Negativas",,, .F. ,,0,16777215,,, .T. ,,, .T.  )

	oFWL:=	FWLayer():New()
	oFWL:Init(oDlg2, .F. )
	oFWL:AddLine("L1",100, .T. )
	oFWL:AddCollumn("C1",08, .F. ,"L1")
	oFWL:AddCollumn("C2",92, .F. ,"L1")
	oFWL:AddWindow("C1","J1","Rotinas"	,100, .T. , .T. ,,"L1",{||})
	oFWL:AddWindow("C2","J1","Resumo"	,100, .T. , .T. ,,"L1",{||})

	oPan1:=	oFWL:GetWinPanel("C1","J1","L1")
	oPan2:= oFWL:GetWinPanel("C2","J1","L1")


	oBtn01	:=	TButton():New((aPosObj[1,3]/100)*05	,(aPosObj[1,4]/100)*0.5,"Exp.Excel"	,oPan1,{||MsgRun("Atualizando dados...","Processando....",{||PRCEXCEL()})}	,(aPosObj[1,3]/100)*12,10,,oFont1, .F. , .T. , .F. ,, .F. ,,, .F. )
	oBtn02	:=	TButton():New((aPosObj[1,3]/100)*10	,(aPosObj[1,4]/100)*0.5,"Processar"	,oPan1,{||PrcOp01A()}														,(aPosObj[1,3]/100)*12,10,,oFont1, .F. , .T. , .F. ,, .F. ,,, .F. )
	oBtn03	:=	TButton():New((aPosObj[1,3]/100)*90	,(aPosObj[1,4]/100)*0.5,"Sair"		,oPan1,{||oDlg2:End()}														,(aPosObj[1,3]/100)*12,10,,oFont1, .F. , .T. , .F. ,, .F. ,,, .F. )

	oGetDados01 := MsNewGetDados():New( aPosObj[1,1], aPosObj[1,2], aPosObj[1,3], aPosObj[1,4], 2, "AllwaysTrue", "AllwaysTrue", "AllwaysTrue",{},0, 999, "AllwaysTrue", "", "AllwaysTrue", oPan2, aCab01, aItens01)
	oGetDados01:oBrowse:Align		:= 5
	oGetDados01:oBrowse:bLDblClick	:= {||PrcLin()}
	oGetDados01:oBrowse:bHeaderClick:= {||PrcCol(),oGetDados01:Refresh()}

	oGetDados01:oBrowse:lUseDefaultColors := .F.
	oGetDados01:oBrowse:SetBlkBackColor({|| PC02C02B("oGetDados01")})
	oGetDados01:bChange:= {|| PC02C02C("oGetDados01")}

	oDlg2:Activate(,,,,)

Return

Static Function LoadData()

	Local _aAreaSM0 := SM0->(GetArea())
	Local cTop01	:="SQL01"

	lEnd := If( lEnd == nil, .F. , lEnd )

	If Type("oGetDados01") <> "O"

		Aadd(aCab01, {""				,"COR1"		,"@BMP"							,2							,0							, .T. 		,				,""	,"",""})
		Aadd(aCab01, {"Filial"			,"FILIAL"	,PesqPict("SC2","C2_FILIAL"	)	,TamSx3("C2_FILIAL"	)	[1]	,TamSx3("C2_FILIAL"	)	[2]	,""			,"","C",""		,"R","","",""})
		Aadd(aCab01, {"Codigo"			,"PRODUTO"	,PesqPict("SC2","C2_PRODUTO")	,TamSx3("C2_PRODUTO")	[1]	,TamSx3("C2_PRODUTO")	[2]	,""			,"","C",""		,"R","","",""})
		Aadd(aCab01, {"Descrição"		,"DESCRI"	,PesqPict("SB1","B1_DESC"	)	,TamSx3("B1_DESC"	)	[1]	,TamSx3("B1_DESC"	)	[2]	,""			,"","C",""		,"R","","",""})
		Aadd(aCab01, {"Tipo"			,"DESCRI"	,PesqPict("SB1","B1_TIPO"	)	,TamSx3("B1_TIPO"	)	[1]	,TamSx3("B1_TIPO"	)	[2]	,""			,"","C",""		,"R","","",""})
		Aadd(aCab01, {"Armazem"			,"ARMAZEM"	,PesqPict("SC2","C2_LOCAL"	)	,TamSx3("C2_LOCAL"	)	[1]	,TamSx3("C2_LOCAL"	)	[2]	,""			,"","C",""		,"R","","",""})
		Aadd(aCab01, {"Ult.Fecha"		,"DTFECHA"	,PesqPict("SC2","C2_EMISSAO")	,TamSx3("C2_EMISSAO")	[1]	,TamSx3("C2_EMISSAO")	[2]	,""			,"","D",""		,"R","","",""})
		Aadd(aCab01, {"Dt.Inicial"		,"DTINI"	,PesqPict("SC2","C2_EMISSAO")	,TamSx3("C2_EMISSAO")	[1]	,TamSx3("C2_EMISSAO")	[2]	,""			,"","D",""		,"R","","",""})
		Aadd(aCab01, {"Dt.Final"		,"DTFIM"	,PesqPict("SC2","C2_EMISSAO")	,TamSx3("C2_EMISSAO")	[1]	,TamSx3("C2_EMISSAO")	[2]	,""			,"","D",""		,"R","","",""})
		Aadd(aCab01, {"Inicial"			,"QTDINI"	,PesqPict("SB2","B2_CM1"	)	,TamSx3("B2_CM1"	)	[1]	,TamSx3("B2_CM1"	)	[2]	,""			,"","N",""		,"R","","",""})
		Aadd(aCab01, {"Entradas"		,"QTDENT"	,PesqPict("SB2","B2_CM1"	)	,TamSx3("B2_CM1"	)	[1]	,TamSx3("B2_CM1"	)	[2]	,""			,"","N",""		,"R","","",""})
		Aadd(aCab01, {"Saidas"			,"QTDSAI"	,PesqPict("SB2","B2_CM1"	)	,TamSx3("B2_CM1"	)	[1]	,TamSx3("B2_CM1"	)	[2]	,""			,"","N",""		,"R","","",""})
		Aadd(aCab01, {"Saldo"			,"SALDO"	,PesqPict("SB2","B2_CM1"	)	,TamSx3("B2_CM1"	)	[1]	,TamSx3("B2_CM1"	)	[2]	,""			,"","N",""		,"R","","",""})

	EndIf

	cFilAntBk	:=	cFilAnt

	aInfEmp		:=	aProcFil
	aItens01	:={}

	nCountSm0 	:= Len(aInfEmp)
	oProcess:SetRegua1(nCountSm0)

	For nX:=1 To Len(aInfEmp)

		SM0->(DbSeek(SM0->M0_CODIGO + aInfEmp[nX,1]))

		oProcess:IncRegua1("Filial processada:" + aInfEmp[nX,1])

		cFilAnt	:=	aInfEmp[nX,1]

		cDtFecha	:= Dtos(GetMv("MV_ULMES"))
		cDtInicio	:= Dtos(GetMv("MV_ULMES")+1)
		cDtFim		:= Dtos(LastDay(GetMv("MV_ULMES")+1))


		cQuery:= " 	SELECT FILIAL,PRODUTO,B1_DESC DESCRI,B1_TIPO TIPO,ARMAZEM,SB9,SD1,D3E,SD2,D3S,SALDO SALDO " + Chr(13)+Chr(10)
		cQuery+= "	FROM (	SELECT FILIAL,PRODUTO,ARMAZEM,SUM(SB9) SB9,SUM(SD1) SD1,SUM(D3E) D3E,SUM(SD2) SD2, ROUND(SUM(D3S),4) D3S, " + Chr(13)+Chr(10)
		cQuery+= " 			ROUND(( SUM(SB9) + SUM(SD1) + SUM(D3E) ) - ( SUM(SD2) + SUM(D3S) ),4) SALDO " + Chr(13)+Chr(10)
		cQuery+= " 			FROM (	SELECT B9_FILIAL FILIAL, B9_COD PRODUTO, B9_LOCAL ARMAZEM,B9_QINI SB9,0 SD1, 0 D3E, 0 SD2,0 D3S " + Chr(13)+Chr(10)
		cQuery+= " 			 		FROM "+RetSqlName("SB9")+" (NOLOCK) SB9 " + Chr(13)+Chr(10)
		cQuery+= " 			 		INNER JOIN "+RetSqlName("SB1")+" (NOLOCK) SB1 ON B1_FILIAL='" +xFilial("SB1")+ "' AND B1_COD=B9_COD AND B1_LOCPAD=B9_LOCAL AND SB1.D_E_L_E_T_!='*' " + Chr(13)+Chr(10)

		cQuery+= " 			 		WHERE B9_FILIAL= '" +xFilial("SB9")+ "' AND B9_DATA='"+cDtFecha+"' AND SB9.D_E_L_E_T_!='*' " + Chr(13)+Chr(10)

		cQuery+= " 					UNION ALL " + Chr(13)+Chr(10)

		cQuery+= " 					SELECT D1_FILIAL FILIAL, D1_COD PRODUTO, MAX(D1_LOCAL) ARMAZEM, 0 SB9, SUM(D1_QUANT) SD1, 0 D3E, 0 SD2,0 D3S " + Chr(13)+Chr(10)
		cQuery+= " 			 		FROM "+RetSqlName("SD1")+" (NOLOCK) SD1 " + Chr(13)+Chr(10)
		cQuery+= " 			 		INNER JOIN "+RetSqlName("SB1")+" (NOLOCK) SB1 ON B1_FILIAL='" +xFilial("SB1")+ "' AND B1_COD=D1_COD AND B1_LOCPAD=D1_LOCAL AND SB1.D_E_L_E_T_!='*' " + Chr(13)+Chr(10)
		cQuery+= " 			 		INNER JOIN "+RetSqlName("SF4")+" (NOLOCK) SF4 ON F4_FILIAL='" +xFilial("SF4")+ "' AND F4_CODIGO=D1_TES AND F4_ESTOQUE='S' AND SF4.D_E_L_E_T_!='*' " + Chr(13)+Chr(10)

		cQuery+= " 			 		WHERE D1_FILIAL= '" +xFilial("SD1")+ "' AND D1_DTDIGIT BETWEEN '" + cDtInicio + "' AND '"+ cDtFim +"'  AND SD1.D_E_L_E_T_!='*' " + Chr(13)+Chr(10)
		cQuery+= " 			 		GROUP BY D1_FILIAL,D1_COD " + Chr(13)+Chr(10)

		cQuery+= " 					UNION ALL " + Chr(13)+Chr(10)

		cQuery+= " 					SELECT D3_FILIAL FILIAL, D3_COD PRODUTO, MAX(D3_LOCAL) ARMAZEM, 0 SB9,0 SD1, SUM(D3_QUANT) D3E, 0 SD2,0 D3S " + Chr(13)+Chr(10)
		cQuery+= " 			 		FROM "+RetSqlName("SD3")+" SD3 " + Chr(13)+Chr(10)
		cQuery+= " 			 		INNER JOIN "+RetSqlName("SB1")+" (NOLOCK) SB1 ON B1_FILIAL='" +xFilial("SB1")+ "' AND B1_COD=D3_COD AND B1_LOCPAD=D3_LOCAL AND SB1.D_E_L_E_T_!='*' " + Chr(13)+Chr(10)

		cQuery+= " 			 		WHERE D3_FILIAL= '" +xFilial("SD3")+ "' AND D3_EMISSAO BETWEEN '" + cDtInicio + "' AND '"+ cDtFim +"' AND D3_TM<'500'  AND D3_ESTORNO=' ' AND SD3.D_E_L_E_T_!='*' " + Chr(13)+Chr(10)
		cQuery+= " 			 		GROUP BY D3_FILIAL,D3_COD " + Chr(13)+Chr(10)

		cQuery+= " 					UNION ALL " + Chr(13)+Chr(10)

		cQuery+= " 					SELECT D2_FILIAL FILIAL, D2_COD PRODUTO, MAX(D2_LOCAL) ARMAZEM, 0 SB9,0 SD1, 0 D3E, SUM(D2_QUANT) SD2,0 D3S " + Chr(13)+Chr(10)
		cQuery+= " 			 		FROM "+RetSqlName("SD2")+" (NOLOCK) SD2 " + Chr(13)+Chr(10)
		cQuery+= " 			 		INNER JOIN "+RetSqlName("SB1")+" (NOLOCK) SB1 ON B1_FILIAL='" +xFilial("SB1")+ "' AND B1_COD=D2_COD AND B1_LOCPAD=D2_LOCAL AND SB1.D_E_L_E_T_!='*' " + Chr(13)+Chr(10)
		cQuery+= " 			 		INNER JOIN "+RetSqlName("SF4")+" (NOLOCK) SF4 ON F4_FILIAL='" +xFilial("SF4")+ "' AND F4_CODIGO=D2_TES AND F4_ESTOQUE='S' AND SF4.D_E_L_E_T_!='*' " + Chr(13)+Chr(10)

		cQuery+= " 			 		WHERE D2_FILIAL= '" +xFilial("SD2")+ "' AND D2_EMISSAO BETWEEN '" + cDtInicio + "' AND '"+ cDtFim +"'   AND SD2.D_E_L_E_T_!='*' " + Chr(13)+Chr(10)
		cQuery+= " 			 		GROUP BY D2_FILIAL,D2_COD " + Chr(13)+Chr(10)

		cQuery+= " 					UNION ALL " + Chr(13)+Chr(10)

		cQuery+= " 					SELECT D3_FILIAL FILIAL, D3_COD PRODUTO, MAX(D3_LOCAL) ARMAZEM, 0 SB9,0 SD1, 0 D3E, 0 SD2,SUM(D3_QUANT) D3S " + Chr(13)+Chr(10)
		cQuery+= " 			 		FROM "+RetSqlName("SD3")+" (NOLOCK) SD3 " + Chr(13)+Chr(10)
		cQuery+= " 			 		INNER JOIN "+RetSqlName("SB1")+" (NOLOCK) SB1 ON B1_FILIAL='" +xFilial("SB1")+ "' AND B1_COD=D3_COD AND B1_LOCPAD=D3_LOCAL AND SB1.D_E_L_E_T_!='*' " + Chr(13)+Chr(10)

		cQuery+= " 			 		WHERE D3_FILIAL= '" +xFilial("SD3")+ "' AND D3_EMISSAO BETWEEN '" + cDtInicio + "' AND '"+ cDtFim +"' AND D3_TM>'500' AND D3_ESTORNO=' ' AND SD3.D_E_L_E_T_!='*' " + Chr(13)+Chr(10)
		cQuery+= " 			 		GROUP BY D3_FILIAL,D3_COD	) TMP1 " + Chr(13)+Chr(10)

		cQuery+= " 			GROUP BY FILIAL, PRODUTO, ARMAZEM	)TMP2 " + Chr(13)+Chr(10)
		cQuery+= " INNER JOIN "+RetSqlName("SB1")+" (NOLOCK) SB1 ON B1_FILIAL='" + xFilial("SB1") + "' AND B1_COD=PRODUTO AND SB1.D_E_L_E_T_!='*' " + Chr(13)+Chr(10)
		cQuery+= " WHERE SALDO<0 " + Chr(13)+Chr(10)
		cQuery+= " ORDER BY 1,2 " + Chr(13)+Chr(10)

		If !Empty(Select(cTop01))
			DbSelectArea(cTop01)
			(cTop01)->(dbCloseArea())
		Endif

		dbUseArea( .T. ,"TOPCONN", TcGenQry( ,,cQuery),cTop01, .T. , .T.  )

		nCount := 0; DBEval( {|| nCount := nCount + 1},,,,,.F. )
		oProcess:SetRegua2(nCount)

		(cTop01)->(dbGoTop())

		If (cTop01)->(!Eof())

			While (cTop01)->(!Eof())

				Aadd(aItens01,{ oLbNo							, (cTop01)->FILIAL				, (cTop01)->PRODUTO				, (cTop01)->DESCRI				, (cTop01)->TIPO					, (cTop01)->ARMAZEM				, Stod(cDtFecha)					, Stod(cDtInicio)					, Stod(cDtFim)					, (cTop01)->SB9					, (cTop01)->SD1 + (cTop01)->D3E	, (cTop01)->SD2 + (cTop01)->D3S	, (cTop01)->SALDO					, .F. })

				(cTop01)->(dbSkip())

				cMsgReg2:=	"Produto: "+AllTrim((cTop01)->PRODUTO)+" - "+AllTRim((cTop01)->DESCRI)

				oProcess:IncRegua2(cMsgReg2)
				//Sleep(200)

			EndDo

		Else

			oProcess:IncRegua2("Sem produtos negativos...")
			Sleep(500)

		EndIf

	Next

	cFilAnt	:=	cFilAntBk

	If Type("oGetDados01") == "O"
		oGetDados01:SetArray(aItens01, .T. )
		oGetDados01:Refresh( .T. )
	EndIf

	RestArea(_aAreaSM0)

Return()

Static Function PC02C02B(oGrid)

	Local nCor1 := 16777215

	Local nCor2 := ( 112 + ( 219 * 256 ) + ( 147 * 65536 ) )
	Local nCor3 := 12632256
	Local nRet  := nCor1

	If &(oGrid):aCols[&(oGrid):nAt][Len(&(oGrid):aCols[&(oGrid):nAt])]
		nRet := nCor3
	ElseIf &(oGrid):nAt = nAtGrid3
		nRet := nCor2
	EndIf

Return(nRet)

Static Function PC02C02C(oGrid)

	nAtGrid3 := &(oGrid):nAt
	&(oGrid):Refresh()

Return

Static Function PrcLin()

	If oGetDados01:oBrowse:ColPos = 1
		If oGetDados01:aCols[oGetDados01:nAt][oGetDados01:oBrowse:COLPOS]:CNAME == "LBNO"
			oGetDados01:aCols[oGetDados01:nAt][oGetDados01:oBrowse:COLPOS]:= oLbOK
		Else
			oGetDados01:aCols[oGetDados01:nAt][oGetDados01:oBrowse:COLPOS]:= oLbNo
		EndIf
		oGetDados01:Refresh()
	EndIf

Return()

Static Function PrcCol()

	If lMark

		If oGetDados01:oBrowse:ColPos = 1

			For nX:=1 To Len(oGetDados01:aCols)
				If oGetDados01:aCols[nX,1]:CNAME == "LBNO"
					oGetDados01:aCols[nX,1]:= oLbOK
				Else
					oGetDados01:aCols[nX,1]:= oLbNo
				EndIf
			Next

		EndIf

		lMark:= .F.

	Else
		lMark:= .T.
	EndIf

Return()

Static Function ProcFil()

	Local aAreaAll	:=	GetArea()
	Local aInfEmp 	:=	FWLoadSM0()
	Local aRetEmp	:=	{}
	Local nOpc		:= 0

	For nX:=1 To Len(aInfEmp)

		If SubStr(aInfEmp[nX,18],1,08) = '27665906'

			If aInfEmp[nX,01] = cEmpAnt
				If SX6->( dbSeek( aInfEmp[nX,02] + "MV_ULMES" )) .And. aInfEmp[nX,02] > "2550"
					Aadd(aProcFil,{ .F. ,aInfEmp[nX,02],AllTrim(aInfEmp[nX,07]),Dtoc(Stod(SX6->X6_CONTEUD))})
				EndIf
			EndIf

		Else

			If aInfEmp[nX,01] = cEmpAnt
				If SX6->( dbSeek( aInfEmp[nX,02] + "MV_ULMES" ))
					Aadd(aProcFil,{ .F. ,aInfEmp[nX,02],AllTrim(aInfEmp[nX,07]),Dtoc(Stod(SX6->X6_CONTEUD))})
				EndIf
			EndIf

		End

	Next

	oDlg1 := MsDialog():New( aCoors[7],aCoors[7], aCoors[6], (aCoors[5]/100)*60, "SELEÇÃO DAS FILIAIS",,, .F. ,,0,16777215,,, .T. ,,, .T.  )

	oBwr01 := FWBrowse():New(oDlg1); If.T.; oBwr01:SetDataArray(.T.); EndIf; If.F.; oBwr01:SetDataText(.F.); EndIf; If.F.; oBwr01:SetDataQuery(.F.); EndIf; If.F.; oBwr01:SetDataTable(.F.); EndIf; If.F.; oBwr01:SetShowLimit(.F.); EndIf; If( ValType() == "C" ); oBwr01:SetAlias(); EndIf; If( ValType() == "B" ); oBwr01:SetDoubleClick(); EndIf; If( ValType() == "B" ); oBwr01:SetLineOk(); EndIf; If( ValType() == "B" ); oBwr01:SetChange(); EndIf; If( ValType() == "B" ); oBwr01:SetAllOK(); EndIf; If( ValType() == "B" ); oBwr01:SetDelete( .T. ,); EndIf; If( ValType() == "B" ); oBwr01:SetDelOK(); EndIf; If( ValType() == "B" ); oBwr01:SetSuperDel(); EndIf; If( ValType() == "B" ); oBwr01:SetEditCell( .T. ,); EndIf; If( ValType() == "B" ); oBwr01:SetGroup(); EndIf; If.F.; oBwr01:SetInsert(.F.); EndIf; If !.F. .And.  (.T. .Or. .F. .Or. .F. .Or. .F.); oBwr01:SetLocate(); EndIf; If !.F.; oBwr01:SetSeek(,); EndIf; If.F.; oBwr01:DisableConfig(); EndIf; If.F.; oBwr01:DisableReport(); EndIf; If.F.; oBwr01:DisableSaveConfig(); EndIf; If( ValType() == "N" ); oBwr01:SetForeColor(); EndIf; If( ValType() == "N" ); oBwr01:SetBackColor(); EndIf; If( ValType() == "N" ); oBwr01:SetClrAlterRow(); EndIf; If( ValType() == "O" ); oBwr01:SetFontBrowse(); EndIf; If( ValType() == "N" ); oBwr01:SetLineHeight(); EndIf; If( ValType() == "N" ); oBwr01:SetLineBegin(); EndIf; If( ValType(aProcFil) == "A" ); oBwr01:SetArray(aProcFil); EndIf; If( ValType() == "C" ); oBwr01:SetFile(); EndIf; If( ValType() == "A" ); oBwr01:SetQueryIndex(); EndIf; If.F.; oBwr01:SetShowLimit(.F.); EndIf; If( ValType() == "C" ); oBwr01:SetQuery(); EndIf; If( ValType() == "A" ); oBwr01:SetFieldFilter(); oBwr01:SetUseFilter(); ElseIf.F.; oBwr01:SetUseFilter(); EndIf; If( ValType() == "C" ); oBwr01:SetFilterDefault(); EndIf; If( ValType() == "C" .And.  ValType() == "C" .And.  ValType() == "C" ); oBwr01:SetFilter(,,); EndIf; If( ValType() == "C" ); oBwr01:SetDescription(); EndIf; If( ValType(oDlg1) == "O" ); oBwr01:SetOwner(oDlg1); EndIf; If.F.; oBwr01:SetNumberLegend(.F.); EndIf; If( ValType() == "C" ); oBwr01:SetProfileID(); EndIf

	oBwr01:DisableConfig()
	oBwr01:DisableLocate()
	oBwr01:DisableReport()


	oBwr01:AddMarkColumns({||If(aProcFil[oBwr01:At()][1],"LBOK","LBNO")}, {||If(aProcFil[oBwr01:At()][1],aProcFil[oBwr01:At()][1]:= .F. ,aProcFil[oBwr01:At()][1]:= .T. ),oBwr01:Refresh()}, {||MsgRun("Marcando registros...","Processando....",{||PrcAll()})})
	oColumn := FWBrwColumn():New(); If( ValType(1) == "N" ); oColumn:SetAlign(1); EndIf; If( ValType() == "N" ); oColumn:SetBackColor(); EndIf; If( ValType() == "C" ); oColumn:SetComment(); EndIf; If( ValType({||aProcFil[oBwr01:At()][2]}) == "B" ); oColumn:SetData({||aProcFil[oBwr01:At()][2]}); EndIf; If( ValType(TamSx3("A1_FILIAL")[2]) == "N" ); oColumn:SetDecimal(TamSx3("A1_FILIAL")[2]); EndIf; If.F.; oColumn:SetDelete(.F.); EndIf; If.F.; oColumn:SetDetails(.F.); EndIf; If( ValType() == "B" ); oColumn:SetDoubleClick(); EndIf; If.F.; oColumn:SetEdit(.F.); EndIf; If( ValType() == "N" ); oColumn:SetForeColor(); EndIf; If( ValType() == "B" ); oColumn:SetHeaderClick(); EndIf; oColumn:SetOptions(); If( ValType() == "N" ); oColumn:SetOrder(); EndIf; If( ValType() == "C"); oColumn:SetPicture(); ElseIf( ValType() == "B"); oColumn:SetPicture(); EndIf; If( ValType() == "C" ); oColumn:SetReadVar(); EndIf; If( ValType(TamSx3("A1_FILIAL")[1]) == "N" ); oColumn:SetSize(TamSx3("A1_FILIAL")[1]); EndIf; If(.F. ); oColumn:SetAutoSize( .F. ); EndIf; If.F.; oColumn:SetImage(.F.); EndIf; If( ValType("Filial") == "C" ); oColumn:SetTitle("Filial"); Else; oColumn:SetTitle("  "); EndIf; If( ValType() == "C" ); oColumn:SetType(); EndIf; If( ValType() == "B" ); oColumn:SetValid(); EndIf; If( ValType() == "C" ); oColumn:SetID(); EndIf; If( ValType(oBwr01) == "O" ); If ( oBwr01:ClassName() $ "FWBROWSE|FWFORMBROWSE|FWMBROWSE|FWMARKBROWSE" ); oBwr01:SetColumns({oColumn}); EndIf; EndIf
	oColumn := FWBrwColumn():New(); If( ValType(1) == "N" ); oColumn:SetAlign(1); EndIf; If( ValType() == "N" ); oColumn:SetBackColor(); EndIf; If( ValType() == "C" ); oColumn:SetComment(); EndIf; If( ValType({||aProcFil[oBwr01:At()][3]}) == "B" ); oColumn:SetData({||aProcFil[oBwr01:At()][3]}); EndIf; If( ValType(TamSx3("A1_NOME")[2]) == "N" ); oColumn:SetDecimal(TamSx3("A1_NOME")[2]); EndIf; If.F.; oColumn:SetDelete(.F.); EndIf; If.F.; oColumn:SetDetails(.F.); EndIf; If( ValType() == "B" ); oColumn:SetDoubleClick(); EndIf; If.F.; oColumn:SetEdit(.F.); EndIf; If( ValType() == "N" ); oColumn:SetForeColor(); EndIf; If( ValType() == "B" ); oColumn:SetHeaderClick(); EndIf; oColumn:SetOptions(); If( ValType() == "N" ); oColumn:SetOrder(); EndIf; If( ValType() == "C"); oColumn:SetPicture(); ElseIf( ValType() == "B"); oColumn:SetPicture(); EndIf; If( ValType() == "C" ); oColumn:SetReadVar(); EndIf; If( ValType(TamSx3("A1_NOME")[1]) == "N" ); oColumn:SetSize(TamSx3("A1_NOME")[1]); EndIf; If(.F. ); oColumn:SetAutoSize( .F. ); EndIf; If.F.; oColumn:SetImage(.F.); EndIf; If( ValType("Descrição") == "C" ); oColumn:SetTitle("Descrição"); Else; oColumn:SetTitle("  "); EndIf; If( ValType() == "C" ); oColumn:SetType(); EndIf; If( ValType() == "B" ); oColumn:SetValid(); EndIf; If( ValType() == "C" ); oColumn:SetID(); EndIf; If( ValType(oBwr01) == "O" ); If ( oBwr01:ClassName() $ "FWBROWSE|FWFORMBROWSE|FWMBROWSE|FWMARKBROWSE" ); oBwr01:SetColumns({oColumn}); EndIf; EndIf
	oColumn := FWBrwColumn():New(); If( ValType(1) == "N" ); oColumn:SetAlign(1); EndIf; If( ValType() == "N" ); oColumn:SetBackColor(); EndIf; If( ValType() == "C" ); oColumn:SetComment(); EndIf; If( ValType({||aProcFil[oBwr01:At()][4]}) == "B" ); oColumn:SetData({||aProcFil[oBwr01:At()][4]}); EndIf; If( ValType(0) == "N" ); oColumn:SetDecimal(0); EndIf; If.F.; oColumn:SetDelete(.F.); EndIf; If.F.; oColumn:SetDetails(.F.); EndIf; If( ValType() == "B" ); oColumn:SetDoubleClick(); EndIf; If.F.; oColumn:SetEdit(.F.); EndIf; If( ValType() == "N" ); oColumn:SetForeColor(); EndIf; If( ValType() == "B" ); oColumn:SetHeaderClick(); EndIf; oColumn:SetOptions(); If( ValType() == "N" ); oColumn:SetOrder(); EndIf; If( ValType() == "C"); oColumn:SetPicture(); ElseIf( ValType() == "B"); oColumn:SetPicture(); EndIf; If( ValType() == "C" ); oColumn:SetReadVar(); EndIf; If( ValType(10) == "N" ); oColumn:SetSize(10); EndIf; If(.F. ); oColumn:SetAutoSize( .F. ); EndIf; If.F.; oColumn:SetImage(.F.); EndIf; If( ValType("Dt.Fechamento") == "C" ); oColumn:SetTitle("Dt.Fechamento"); Else; oColumn:SetTitle("  "); EndIf; If( ValType() == "C" ); oColumn:SetType(); EndIf; If( ValType() == "B" ); oColumn:SetValid(); EndIf; If( ValType() == "C" ); oColumn:SetID(); EndIf; If( ValType(oBwr01) == "O" ); If ( oBwr01:ClassName() $ "FWBROWSE|FWFORMBROWSE|FWMBROWSE|FWMARKBROWSE" ); oBwr01:SetColumns({oColumn}); EndIf; EndIf

	oBwr01:Activate()

	oDlg1:Activate(,,,,EnchoiceBar(oDlg1,{|| nOpc:=1,oDlg1:End()},{||oDlg1:End()},,),,)

	If nOpc=1
		For nX:=1 To Len(aProcFil)
			If aProcFil[nX,01]
				Aadd(aRetEmp,{aProcFil[nX,02]})
			EndIf
		Next
	EndIf

	RestArea(aAreaAll)

Return(aRetEmp)

Static Function PrcAll()

	For nX:=1 To Len(aProcFil)
		If aProcFil[nX,01]
			aProcFil[nX,01]	:= .F.
		Else
			aProcFil[nX,01]	:= .T.
		EndIf
	Next

	oBwr01:Refresh()

Return()


Static Function PrcOp01A()

	Local aItOp		:=	{}
	Private cCodMov	:=	GetMv("CODMOVENT1")

	If !Empty(cCodMov)

		For nX:=1 To Len(oGetDados01:aCols)
			If oGetDados01:aCols[nX,1]:CNAME == "LBOK"
				AAdd(aItOp,{ oGetDados01:aCols[nX,02],				oGetDados01:aCols[nX,03],				oGetDados01:aCols[nX,04],				oGetDados01:aCols[nX,06],				oGetDados01:aCols[nX,08],				oGetDados01:aCols[nX,09],				Abs(oGetDados01:aCols[nX,13])				})
			EndIf
		Next

		If Len(aItOp) > 0

			oProcess := MsNewProcess():New({|lEnd| PrcOp02(aItOp,@oProcess, @lEnd) },"EMISSÃO DOS MOVIMENTOS","Preparando ambiente...", .T. )
			oProcess:Activate()

			oProcess := MsNewProcess():New({|lEnd| LoadData(@oProcess, @lEnd) },"Analisando Saldo Pós Movimentação","Acessando banco de dados", .T. )
			oProcess:Activate()

		Else
			Iif(FindFunction("APMsgInfo"), APMsgInfo("Nenhum produto foi selecionado.", ".:ATENÇÃO:."), MsgInfo("Nenhum produto foi selecionado.", ".:ATENÇÃO:."))
		EndIf

	Else

		Iif(FindFunction("APMsgStop"), APMsgStop("Codigo da movimentação não cadastrado parâmetro CODMOVENT1", ".:ERRO:."), MsgStop("Codigo da movimentação não cadastrado parâmetro CODMOVENT1", ".:ERRO:."))

	EndIf

Return()

Static Function PrcOp02(aItOp)

	Local nReg	:=	Len(aItOp)

	oProcess:SetRegua1(nReg)
	oProcess:SetRegua2(nX)

	For nX:=1 To nReg

		Begin Sequence; BeginTran()

			oProcess:IncRegua1("Filial em processamento :" + aItOp[nX,1])
			oProcess:IncRegua2("Produto: " + AllTrim(aItOp[nX,3]) )
			PrcOp03(aItOp[nX])

			Sleep(500)

		EndTran(); end

	Next

Return()

Static Function PrcOp03(aItens)

	Local aArea03	:=	GetArea()
	Local nOpc		:=	3
	Local cFilPrc	:=	aItens[01]
	Local cProd		:=	aItens[02]
	Local cLocal	:=	aItens[04]
	Local cDtIni	:=	Dtos(aItens[05])
	Local cDtFim	:=	Dtos(aItens[06])
	Local nSaldo	:=	aItens[07]
	Local aItensSD3	:=	{}

	Private lMsErroAuto	:= .F.
	Private lMsHelpAuto := .T.

	dBkDtBase	:= dDataBase
	dDataBase	:= aItens[06]

	SM0->(DbSeek(SM0->M0_CODIGO + aItens[01]))
	cFilAnt	:=	aItens[01]

	aCabSd3	:=	{ {"D3_TM"		,cCodMov	,NIL}, {"D3_EMISSAO"	,ddatabase	,NIL}}

	aItSD3:={ {"D3_COD"		,cProd		,NIL}, {"D3_UM"		,"UN"		,NIL}, {"D3_QUANT"		,nSaldo		,NIL}, {"D3_LOCAL"		,cLocal		,NIL}, {"D3_OBSERVA"	,"AUTMPEM/Dt Proc:"+Dtoc(Date())	,NIL}}

	aadd(aItensSD3,aItSD3)

	MSExecAuto({|x,y,z| MATA241(x,y,z)},aCabSd3,aItensSD3,3)

	If lMsErroAuto
		Alert(MostraErro())
		DisarmTransaction()
	Else
		lMsErroAuto	:= .F.
	EndIf

	dDataBase:=	dBkDtBase

	RestArea(aArea03)

Return()

Static Function PRCEXCEL()

	Private cPath	:=	""
	Private cSheet1	:=	""
	Private cTable	:=	"Resumo movimentação"
	Private oExcel	:=	FWMSEXCEL():New()

	If !ExistDir("C:\TEMP")
		MakeDir("C:\TEMP")
	EndIf

	cPath:= Iif(FindFunction("FWHasAccMode") .And.  FindFunction("AVGetFile") .And.  FWHasAccMode(1), AVGetFile("Arquivos xls  (*.xls)  | *.xls  ", " ", 1, "C:\TEMP", .T. , 48+128, .F. , .T. ,,,,,), cGetFile("Arquivos xls  (*.xls)  | *.xls  ", " ", 1, "C:\TEMP", .T. , 48+128,, .T. ))

	If Empty(cPath)
		Iif(FindFunction("APMsgStop"), APMsgStop("Diretório incorreto!!!", "Erro"), MsgStop("Diretório incorreto!!!", "Erro"))
		Return()
	Else

		cFilProc:=""

		For nX:=1 to Len(aItens01)

			If aItens01[nX,02] <> cFilProc

				cSheet1 := " Filial-" + aItens01[nX,02]

				oExcel:AddworkSheet(cSheet1)

				oExcel:AddTable (cSheet1,cTable)
				oExcel:AddColumn(cSheet1,cTable,"FILIAL"			,1,1)
				oExcel:AddColumn(cSheet1,cTable,"CODIGO"			,1,1)
				oExcel:AddColumn(cSheet1,cTable,"DESCRIÇÃO"			,1,1)
				oExcel:AddColumn(cSheet1,cTable,"TIPO"				,1,1)
				oExcel:AddColumn(cSheet1,cTable,"ARMAZEM"			,1,1)
				oExcel:AddColumn(cSheet1,cTable,"ULT.FECHA"			,1,1)
				oExcel:AddColumn(cSheet1,cTable,"DT.INICIAL"		,1,1)
				oExcel:AddColumn(cSheet1,cTable,"DT.FINAL"			,1,1)
				oExcel:AddColumn(cSheet1,cTable,"INICIAL"			,1,2)
				oExcel:AddColumn(cSheet1,cTable,"ENTRADAS"			,1,2)
				oExcel:AddColumn(cSheet1,cTable,"SAIDAS"			,1,2)
				oExcel:AddColumn(cSheet1,cTable,"SALDO"				,1,2)

			EndIf

			cFilProc	:=	aItens01[nX,02]

			oExcel:AddRow(cSheet1,cTable,{ aItens01[nX,02], aItens01[nX,03], aItens01[nX,04], aItens01[nX,05], aItens01[nX,06], aItens01[nX,07], aItens01[nX,08], aItens01[nX,09], aItens01[nX,10], aItens01[nX,11], aItens01[nX,12], aItens01[nX,13] })

		Next

		cArq:= cPath+"RESMOVMPEM.xls"

		oExcel:Activate()
		oExcel:GetXMLFile(cArq)
		oExcelApp := MsExcel():New()
		oExcelApp:WorkBooks:Open(cArq)
		oExcelApp:SetVisible( .T. )

	EndIf

Return()

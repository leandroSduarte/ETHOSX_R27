#include 'protheus.ch'
#include 'parmtype.ch'
#INCLUDE "FWMBROWSE.CH"
#INCLUDE 'FWMVCDEF.CH'

User Function PrcOp01()

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
	Private lMark		:=	.T.
	Private oFont1		:=	TFont():New('Times New Roman',,-12,.T.)
	Private oFont2		:=	TFont():New('Times New Roman',,-10,.T.)
	Private oProcess


	AAdd(aObjects,{100,100,.T.,.T.})
	aInfo:={aCoors[1],aCoors[2],aCoors[3],aCoors[4],5,5}
	aPosObj:= MsObjSize(aInfo,aObjects,.T.,.F.)

	aProcFil:= ProcFil()

	If Len(aProcFil)=0
		msgStop("Nenhuma filial foi selecionada",".:ERRO:.")
		Return()
	Else
		oProcess := MsNewProcess():New({|lEnd| LoadData(@oProcess, @lEnd) },"Analisando movimentação","Acessando banco de dados",.T.)
		oProcess:Activate()
	EndIf

	oDlg2 := MsDialog():New( aCoors[7],aCoors[7], aCoors[6], aCoors[5], "Painel Emissão Ops x PAs negativos",,,.F.,,CLR_BLACK,CLR_WHITE,,,.T.,,,.T. )

	oFWL:=	FWLayer():New()
	oFWL:Init(oDlg2,.F.)
	oFWL:AddLine("L1",100,.T.)
	oFWL:AddCollumn("C1",08,.F.,"L1")
	oFWL:AddCollumn("C2",92,.F.,"L1")
	oFWL:AddWindow("C1","J1",'Rotinas'	,100,.T.,.T.,,"L1",{||})
	oFWL:AddWindow("C2","J1",'Produto'	,085,.T.,.T.,,"L1",{||})
	//oFWL:AddWindow("C2","J2",'detalhes'	,015,.T.,.T.,,"L1",{||})
	oPan1:=	oFWL:GetWinPanel("C1","J1","L1")
	oPan2:= oFWL:GetWinPanel("C2","J1","L1")
	//oPan3:= oFWL:GetWinPanel("C2","J2","L1")

	oBtn01	:=	TButton():New((aPosObj[1,3]/100)*05	,(aPosObj[1,4]/100)*0.5,"Exp.Excel"	,oPan1,{||MsgRun("Atualizando dados...","Processando....",{||PRCEXCEL()})}	,(aPosObj[1,3]/100)*12,10,,oFont1,.F.,.T.,.F.,,.F.,,,.F.)
	oBtn02	:=	TButton():New((aPosObj[1,3]/100)*10	,(aPosObj[1,4]/100)*0.5,"Processar"	,oPan1,{||PrcOp01A()}															,(aPosObj[1,3]/100)*12,10,,oFont1,.F.,.T.,.F.,,.F.,,,.F.)
	oBtn03	:=	TButton():New((aPosObj[1,3]/100)*90	,(aPosObj[1,4]/100)*0.5,"Sair"		,oPan1,{||oDlg2:End()}														,(aPosObj[1,3]/100)*12,10,,oFont1,.F.,.T.,.F.,,.F.,,,.F.)

	oGetDados01 := MsNewGetDados():New( aPosObj[1,1], aPosObj[1,2], aPosObj[1,3], aPosObj[1,4], GD_UPDATE, "AllwaysTrue", "AllwaysTrue", "AllwaysTrue",{},0, 999, "AllwaysTrue", "", "AllwaysTrue", oPan2, aCab01, aItens01)
	oGetDados01:oBrowse:Align		:= CONTROL_ALIGN_ALLCLIENT
	oGetDados01:oBrowse:bLDblClick	:= {||PrcLin()}
	oGetDados01:oBrowse:bHeaderClick:= {||PrcCol(),oGetDados01:Refresh()}

	oGetDados01:oBrowse:lUseDefaultColors := .F.
	oGetDados01:oBrowse:SetBlkBackColor({|| PC02C02B("oGetDados01")})
	oGetDados01:bChange:= {|| PC02C02C("oGetDados01")}

	oDlg2:Activate(,,,,)

Return

Static Function ProcFil()

	Local aAreaAll	:=	GetArea()
	Local aInfEmp 	:=	FWLoadSM0()
	Local aRetEmp	:=	{}
	Local nOpc		:= 0

	For nX:=1 To Len(aInfEmp)

		If aInfEmp[nX,01] = "20"
			If SX6->( dbSeek( aInfEmp[nX,02] + "MV_ULMES" ))
				Aadd(aProcFil,{.F.,aInfEmp[nX,02],AllTrim(aInfEmp[nX,07]),Dtoc(Stod(SX6->X6_CONTEUD))})
			EndIf
		EndIf

	Next

	oDlg1 := MsDialog():New( aCoors[7],aCoors[7], aCoors[6], (aCoors[5]/100)*60, "SELEÇÃO DAS FILIAIS",,,.F.,,CLR_BLACK,CLR_WHITE,,,.T.,,,.T. )

	DEFINE FWBROWSE oBwr01 DATA ARRAY ARRAY aProcFil OF oDlg1

	oBwr01:DisableConfig()
	oBwr01:DisableLocate()
	oBwr01:DisableReport()	
	oBwr01:AddMarkColumns({||If(aProcFil[oBwr01:At()][1],'LBOK','LBNO')},;
	{||If(aProcFil[oBwr01:At()][1],aProcFil[oBwr01:At()][1]:=.F.,aProcFil[oBwr01:At()][1]:=.T.),oBwr01:Refresh()},;
	{||MsgRun("Marcando registros...","Processando....",{||PrcAll()})})
	ADD COLUMN oColumn    DATA    { || aProcFil[oBwr01:At()][2] }   TITLE "Filial"			SIZE TamSx3("A1_FILIAL")[1]	DECIMAL TamSx3("A1_FILIAL")	[2] ALIGN 1 OF oBwr01
	ADD COLUMN oColumn    DATA    { || aProcFil[oBwr01:At()][3] }   TITLE "Descrição"		SIZE TamSx3("A1_NOME")	[1]	DECIMAL TamSx3("A1_NOME")	[2]	ALIGN 1 OF oBwr01
	ADD COLUMN oColumn    DATA    { || aProcFil[oBwr01:At()][4] }   TITLE "Dt.Fechamento"	SIZE 10						DECIMAL 0						ALIGN 1 OF oBwr01

	ACTIVATE FWBROWSE oBwr01

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
			aProcFil[nX,01]	:=	.F.
		Else
			aProcFil[nX,01]	:=	.T.
		EndIf
	Next

	oBwr01:Refresh()

Return()

Static Function LoadData()

	Local _aAreaSM0 := SM0->(GetArea())
	Local cTop01	:="SQL01"

	Default lEnd := .F.

	If Type("oGetDados01") != "O"

		Aadd(aCab01, {""				,"COR1"		,"@BMP"							,2							,0							,.T.		,				,""	,"",""})
		Aadd(aCab01, {"Filial"			,"FILIAL"	,PesqPict('SC2','C2_FILIAL'	)	,TamSx3('C2_FILIAL'	)	[1]	,TamSx3('C2_FILIAL'	)	[2]	,""			,"","C",""		,"R","","",""})
		Aadd(aCab01, {"Codigo"			,"PRODUTO"	,PesqPict('SC2','C2_PRODUTO')	,TamSx3('C2_PRODUTO')	[1]	,TamSx3('C2_PRODUTO')	[2]	,""			,"","C",""		,"R","","",""})
		Aadd(aCab01, {"Descrição"		,"DESCRI"	,PesqPict('SB1','B1_DESC'	)	,TamSx3('B1_DESC'	)	[1]	,TamSx3('B1_DESC'	)	[2]	,""			,"","C",""		,"R","","",""})
		Aadd(aCab01, {"Tipo"			,"DESCRI"	,PesqPict('SB1','B1_TIPO'	)	,TamSx3('B1_TIPO'	)	[1]	,TamSx3('B1_TIPO'	)	[2]	,""			,"","C",""		,"R","","",""})
		Aadd(aCab01, {"Armazem"			,"ARMAZEM"	,PesqPict('SC2','C2_LOCAL'	)	,TamSx3('C2_LOCAL'	)	[1]	,TamSx3('C2_LOCAL'	)	[2]	,""			,"","C",""		,"R","","",""})
		Aadd(aCab01, {"Ult.Fecha"		,"DTFECHA"	,PesqPict('SC2','C2_EMISSAO')	,TamSx3('C2_EMISSAO')	[1]	,TamSx3('C2_EMISSAO')	[2]	,""			,"","D",""		,"R","","",""})
		Aadd(aCab01, {"Dt.Inicial"		,"DTINI"	,PesqPict('SC2','C2_EMISSAO')	,TamSx3('C2_EMISSAO')	[1]	,TamSx3('C2_EMISSAO')	[2]	,""			,"","D",""		,"R","","",""})
		Aadd(aCab01, {"Dt.Final"		,"DTFIM"	,PesqPict('SC2','C2_EMISSAO')	,TamSx3('C2_EMISSAO')	[1]	,TamSx3('C2_EMISSAO')	[2]	,""			,"","D",""		,"R","","",""})
		Aadd(aCab01, {"Inicial"			,"QTDINI"	,PesqPict('SB2','B2_CM1'	)	,TamSx3('B2_CM1'	)	[1]	,TamSx3('B2_CM1'	)	[2]	,""			,"","N",""		,"R","","",""})
		Aadd(aCab01, {"Entradas"		,"QTDENT"	,PesqPict('SB2','B2_CM1'	)	,TamSx3('B2_CM1'	)	[1]	,TamSx3('B2_CM1'	)	[2]	,""			,"","N",""		,"R","","",""})
		Aadd(aCab01, {"Saidas"			,"QTDSAI"	,PesqPict('SB2','B2_CM1'	)	,TamSx3('B2_CM1'	)	[1]	,TamSx3('B2_CM1'	)	[2]	,""			,"","N",""		,"R","","",""})
		Aadd(aCab01, {"Saldo"			,"SALDO"	,PesqPict('SB2','B2_CM1'	)	,TamSx3('B2_CM1'	)	[1]	,TamSx3('B2_CM1'	)	[2]	,""			,"","N",""		,"R","","",""})

	EndIf

	cFilAntBk	:=	cFilAnt
	//aInfEmp		:=	FWAllFilial()
	aInfEmp		:=	aProcFil
	aItens01	:=	{}

	nCountSm0 	:= Len(aInfEmp)
	oProcess:SetRegua1(nCountSm0)

	For nX:=1 To Len(aInfEmp)

		SM0->(DbSeek(SM0->M0_CODIGO + aInfEmp[nX,1]))

		oProcess:IncRegua1("Filial processada:" + aInfEmp[nX,1])

		cFilAnt	:=	aInfEmp[nX,1]

		cDtFecha	:= Dtos(GetMv("MV_ULMES"))
		cDtInicio	:= Dtos(GetMv("MV_ULMES")+1)
		cDtFim		:= Dtos(LastDay(GetMv("MV_ULMES")+1))

		//cQuery:= " 	SELECT FILIAL,PRODUTO,B1_DESC DESCRI,ARMAZEM,SB9,SD1,D3E,SD2,D3S,ABS(SALDO) SALDO " + CRLF 
		cQuery:= " 	SELECT FILIAL,PRODUTO,B1_DESC DESCRI,B1_TIPO TIPO,ARMAZEM,SB9,SD1,D3E,SD2,D3S,SALDO SALDO " + CRLF 
		cQuery+= "	FROM (	SELECT FILIAL,PRODUTO,ARMAZEM,SUM(SB9) SB9,SUM(SD1) SD1,SUM(D3E) D3E,SUM(SD2) SD2, ROUND(SUM(D3S),4) D3S, " + CRLF
		cQuery+= " 			ROUND(( SUM(SB9) + SUM(SD1) + SUM(D3E) ) - ( SUM(SD2) + SUM(D3S) ),4) SALDO " + CRLF  
		cQuery+= " 			FROM (	SELECT B9_FILIAL FILIAL, B9_COD PRODUTO, B9_LOCAL ARMAZEM,B9_QINI SB9,0 SD1, 0 D3E, 0 SD2,0 D3S " + CRLF 
		cQuery+= " 			 		FROM "+RetSqlName("SB9")+" (NOLOCK) SB9 " + CRLF
		cQuery+= " 			 		INNER JOIN "+RetSqlName("SB1")+" (NOLOCK) SB1 ON B1_FILIAL='" +xFilial("SB1")+ "' AND B1_COD=B9_COD AND B1_TIPO = 'PA' AND SB1.D_E_L_E_T_!='*' " + CRLF 
		cQuery+= " 			 		WHERE B9_FILIAL= '" +xFilial("SB9")+ "' AND B9_DATA='"+cDtFecha+"' AND B9_LOCAL='01' AND SB9.D_E_L_E_T_!='*' " + CRLF  

		cQuery+= " 					UNION ALL " + CRLF  

		cQuery+= " 					SELECT D1_FILIAL FILIAL, D1_COD PRODUTO, MAX(D1_LOCAL) ARMAZEM, 0 SB9, SUM(D1_QUANT) SD1, 0 D3E, 0 SD2,0 D3S " + CRLF 
		cQuery+= " 			 		FROM "+RetSqlName("SD1")+" (NOLOCK) SD1 " + CRLF  
		cQuery+= " 			 		INNER JOIN "+RetSqlName("SF4")+" (NOLOCK) SF4 ON F4_FILIAL='" +xFilial("SF4")+ "' AND F4_CODIGO=D1_TES AND F4_ESTOQUE='S' AND SF4.D_E_L_E_T_!='*' " + CRLF 
		cQuery+= " 			 		WHERE D1_FILIAL= '" +xFilial("SD1")+ "' AND D1_DTDIGIT BETWEEN '" + cDtInicio + "' AND '"+ cDtFim +"' AND D1_LOCAL='01'  AND D1_TP = 'PA' AND SD1.D_E_L_E_T_!='*' " + CRLF   
		cQuery+= " 			 		GROUP BY D1_FILIAL,D1_COD " + CRLF  

		cQuery+= " 					UNION ALL " + CRLF  

		cQuery+= " 					SELECT D3_FILIAL FILIAL, D3_COD PRODUTO, MAX(D3_LOCAL) ARMAZEM, 0 SB9,0 SD1, SUM(D3_QUANT) D3E, 0 SD2,0 D3S " + CRLF 
		cQuery+= " 			 		FROM "+RetSqlName("SD3")+" SD3 " + CRLF   
		cQuery+= " 			 		WHERE D3_FILIAL= '" +xFilial("SD3")+ "' AND D3_EMISSAO BETWEEN '" + cDtInicio + "' AND '"+ cDtFim +"' AND D3_LOCAL='01' AND D3_TM<'500' AND D3_TIPO = 'PA' AND D3_ESTORNO=' ' AND SD3.D_E_L_E_T_!='*' " + CRLF 
		cQuery+= " 			 		GROUP BY D3_FILIAL,D3_COD " + CRLF  

		cQuery+= " 					UNION ALL " + CRLF  

		cQuery+= " 					SELECT D2_FILIAL FILIAL, D2_COD PRODUTO, MAX(D2_LOCAL) ARMAZEM, 0 SB9,0 SD1, 0 D3E, SUM(D2_QUANT) SD2,0 D3S " + CRLF 
		cQuery+= " 			 		FROM "+RetSqlName("SD2")+" (NOLOCK) SD2 " + CRLF  
		cQuery+= " 			 		INNER JOIN "+RetSqlName("SF4")+" (NOLOCK) SF4 ON F4_FILIAL='" +xFilial("SF4")+ "' AND F4_CODIGO=D2_TES AND F4_ESTOQUE='S' AND SF4.D_E_L_E_T_!='*' " + CRLF 
		cQuery+= " 			 		WHERE D2_FILIAL= '" +xFilial("SD2")+ "' AND D2_EMISSAO BETWEEN '" + cDtInicio + "' AND '"+ cDtFim +"' AND D2_LOCAL='01'  AND D2_TP = 'PA'  AND SD2.D_E_L_E_T_!='*' " + CRLF   
		cQuery+= " 			 		GROUP BY D2_FILIAL,D2_COD " + CRLF   

		cQuery+= " 					UNION ALL " + CRLF

		cQuery+= " 					SELECT D3_FILIAL FILIAL, D3_COD PRODUTO, MAX(D3_LOCAL) ARMAZEM, 0 SB9,0 SD1, 0 D3E, 0 SD2,SUM(D3_QUANT) D3S " + CRLF 
		cQuery+= " 			 		FROM "+RetSqlName("SD3")+" (NOLOCK) SD3 " + CRLF   
		cQuery+= " 			 		WHERE D3_FILIAL= '" +xFilial("SD3")+ "' AND D3_EMISSAO BETWEEN '" + cDtInicio + "' AND '"+ cDtFim +"' AND D3_LOCAL='01' AND D3_TM>'500' AND D3_TIPO = 'PA' AND D3_ESTORNO=' ' AND SD3.D_E_L_E_T_!='*' " + CRLF  
		cQuery+= " 			 		GROUP BY D3_FILIAL,D3_COD	) TMP1 " + CRLF  

		cQuery+= " 			GROUP BY FILIAL, PRODUTO, ARMAZEM	)TMP2 " + CRLF  
		cQuery+= " INNER JOIN "+RetSqlName("SB1")+" (NOLOCK) SB1 ON B1_FILIAL='" + xFilial("SB1") + "' AND B1_COD=PRODUTO AND SB1.D_E_L_E_T_!='*' " + CRLF
		cQuery+= " WHERE SALDO<0 " + CRLF 
		cQuery+= " ORDER BY 1,2 " + CRLF

		If !Empty(Select(cTop01))
			DbSelectArea(cTop01)
			(cTop01)->(dbCloseArea())
		Endif

		dbUseArea( .T.,"TOPCONN", TcGenQry( ,,cQuery),cTop01, .T., .T. )

		Count To nCount   
		oProcess:SetRegua2(nCount)

		(cTop01)->(dbGoTop())

		If (cTop01)->(!Eof())

			While (cTop01)->(!Eof())

				Aadd(aItens01,{					;
				oLbNo							,;
				(cTop01)->FILIAL				,;
				(cTop01)->PRODUTO				,;
				(cTop01)->DESCRI				,;
				(cTop01)->TIPO					,;
				(cTop01)->ARMAZEM				,;
				Stod(cDtFecha)					,;
				Stod(cDtInicio)					,;
				Stod(cDtFim)					,;
				(cTop01)->SB9					,;
				(cTop01)->SD1 + (cTop01)->D3E	,;
				(cTop01)->SD2 + (cTop01)->D3S	,;
				(cTop01)->SALDO					,;
				.F.})

				(cTop01)->(dbSkip())

				cMsgReg2:=	"Produto: "+AllTrim((cTop01)->PRODUTO)+" - "+AllTRim((cTop01)->DESCRI)

				oProcess:IncRegua2(cMsgReg2)
				//Sleep(100)

			EndDo

		Else

			oProcess:IncRegua2("Sem produtos negativos...")
			Sleep(300)

		EndIf

	Next

	cFilAnt	:=	cFilAntBk


	If Len(aItens01)=0
		msgInfo("Não foi localizado produtos negativos para as filiais selecionadas!","ATENÇÃO")
	ElseIf Type("oGetDados01") == "O" 
		oGetDados01:SetArray(aItens01,.T.)
		oGetDados01:Refresh(.T.)
	EndIf

	RestArea(_aAreaSM0)

Return()

Static Function PC02C02B(oGrid)

	Local nCor1 := CLR_WHITE  // Branco
	//Local nCor2 := RGB(238,250,95) //CLR_YELLOW// Amarelo
	Local nCor2 := RGB(112,219,147) //CLR_YELLOW// Amarelo
	Local nCor3 := CLR_HGRAY  // Cinza
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

		lMark:=	.F.

	Else
		lMark:=	.T.
	EndIf

Return()

Static Function PrcOp01A()

	Local aItOp:={}

	For nX:=1 To Len(oGetDados01:aCols)
		If oGetDados01:aCols[nX,1]:CNAME == "LBOK"
			AAdd(aItOp,{;
			oGetDados01:aCols[nX,02],;		//FILIAL
			oGetDados01:aCols[nX,03],;		//PRODUTO
			oGetDados01:aCols[nX,04],;		//DESCRIÇÃO
			oGetDados01:aCols[nX,06],;		//ARMAZEM
			oGetDados01:aCols[nX,08],;		//DT INICIAL
			oGetDados01:aCols[nX,09],;		//DT FINAL
			Abs(oGetDados01:aCols[nX,13]);	//SALDO
			})
		EndIf
	Next

	If Len(aItOp) > 0
		oProcess := MsNewProcess():New({|lEnd| PrcOp02(aItOp,@oProcess, @lEnd) },"EMISSÃO OPS INICIADO","Preparando ambiente...",.T.)
		oProcess:Activate()

		oProcess := MsNewProcess():New({|lEnd| LoadData(@oProcess, @lEnd) },"Analisando Saldo Pós Apontamentos","Acessando banco de dados",.T.)
		oProcess:Activate()		

	Else
		msgInfo("Nenhum produto foi selecionado.",".:ATENÇÃO:.")
	EndIf

Return()

Static Function PrcOp02(aItOp)

	Local nReg	:=	Len(aItOp)

	dbSelectArea("SB1")
	SB1->(dbSetOrder(1))

	oProcess:SetRegua1(nReg)
	oProcess:SetRegua2(nX)

	For nX:=1 To nReg

		BEGIN TRANSACTION

			oProcess:IncRegua1("Filial em processamento :" + aItOp[nX,1])

			oProcess:IncRegua2("Verificando OPS em aberto para produto")
			PrcOp04(aItOp[nX])

			oProcess:IncRegua2("Incluindo OP: " + AllTrim(aItOp[nX,3]) )
			PrcOp03(aItOp[nX])

			oProcess:IncRegua2("Apontando OP: " + AllTrim(aItOp[nX,3]) )
			PrcOp04(aItOp[nX])

			//Sleep(500)

		END TRANSACTION

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

	Private lMsErroAuto	:=	.F.

	dBkDtBase	:= dDataBase
	dDataBase	:= aItens[06]

	SM0->(DbSeek(SM0->M0_CODIGO + aItens[01]))
	cFilAnt	:=	aItens[01]

	cNumOP := NumOpPa()

	aItensOp	:={;
	{'C2_FILIAL'	,cFilPrc				,NIL},;
	{'C2_NUM'		,cNumOP					,NIL},;
	{'C2_ITEM'		,"01"					,NIL},;
	{'C2_SEQUEN'	,"001"					,NIL},;
	{'C2_PRODUTO'	,cProd					,NIL},;
	{'C2_LOCAL'		,"01"					,NIL},;
	{'C2_QUANT'		,nSaldo					,NIL},;
	{'C2_UM'		,"UN"					,NIL},;
	{'C2_DATPRI'	,dDataBase				,NIL},;
	{'C2_OBS'		,"AUTOPPA"				,NIL},;
	{'C2_DATPRF'	,dDataBase				,NIL},;
	{'C2_EMISSAO'	,dDataBase				,NIL},;
	{'C2_TPOP'		,"F"					,NIL},;
	{'C2_TPPR'		,"I"					,NIL},;
	{'AUTEXPLODE'	,"S"					,NIL}}

	//ÚÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄ¿
	//³ Se alteracao ou exclusao, deve-se posicionar no registro     ³
	//³ da SC2 antes de executar a rotina automatica                 ³
	//ÀÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÄÙ
	/*If nOpc == 4 .Or. nOpc == 5
	SC2->(DbSetOrder(1)) // FILIAL + NUM + ITEM + SEQUEN + ITEMGRD
	SC2->(DbSeek(xFilial("SC2")+"000097"+"01"+"002"))
	EndIf*/

	MsExecAuto({|x,Y| Mata650(x,Y)},aItensOp,nOpc)

	If lMsErroAuto
		Alert(MostraErro())
		//ConOut("AutOpPa - Emp: "+aEmpPrc[1]+" - Filial: "+aEmpPrc[2]+" - Erro para gerar OP produto:"+AllTrim((cTop01)->PRODUTO)+" - "+dtoc(date())+" - "+time())
		//RollBackSx8()
		DisarmTransaction()
	Else
		//ConOut("AutOpPa - Emp: "+aEmpPrc[1]+" - Filial: "+aEmpPrc[2]+" - OP: "+cNumOP+" gerada produto:"+AllTrim((cTop01)->PRODUTO)+" - Dt.Op "+dtoc(dDataBase)+" - "+time())
		//ConfirmSX8()
		lMsErroAuto	:=	.F.
		//cNumOP:= Soma1(cNumOP)
		//PutMv('NUMAUTOP',cNumOP)
	EndIf

	dDataBase:=	dBkDtBase

	RestArea(aArea03)

Return()

Static Function PrcOp04(aItens)

	Local aArea04		:=	GetArea()
	Local cFilPrc	:=	aItens[01]
	Local cProd		:=	aItens[02]
	Local cLocal	:=	aItens[04]
	Local cDtIni	:=	Dtos(aItens[05])
	Local cDtFim	:=	Dtos(aItens[06])
	Local nSaldo	:=	aItens[07]

	Private lMsErroAuto	:=	.F.

	dBkDtBase	:= dDataBase
	dDataBase	:= aItens[06]

	SM0->(DbSeek(SM0->M0_CODIGO + aItens[01]))
	cFilAnt	:=	aItens[01]

	cTM			:= SuperGetMv("TMBXOP",.F.,"001")//CODIGO DA BAIXA
	cTop02		:= "SQL2"

	//cQuery:= " SELECT C2_FILIAL,C2_NUM+'01001' OP, C2_PRODUTO PRODUTO,C2_QUANT SALDO,C2_UM UM,C2_LOCAL ARMAZEM " + CRLF 
	cQuery:= " SELECT C2_FILIAL,C2_NUM+C2_ITEM+C2_SEQUEN OP, C2_PRODUTO PRODUTO,C2_QUANT SALDO,C2_UM UM,C2_LOCAL ARMAZEM " + CRLF 
	cQuery+= " FROM " + RetSqlName("SC2") + " (NOLOCK) SC2  " + CRLF
	cQuery+= " WHERE C2_FILIAL='" + xFilial("SC2") + "' AND C2_PRODUTO='"+cProd+"' AND  C2_EMISSAO BETWEEN '" + cDtIni + "' AND '" + cDtFim + "' " + CRLF
	cQuery+= " AND C2_LOCAL='" + cLocal + "' AND C2_QUANT>C2_QUJE AND C2_OBS='AUTOPPA' AND SC2.D_E_L_E_T_!='*' " + CRLF

	If !Empty(Select(cTop02))
		DbSelectArea(cTop02)
		(cTop02)->(dbCloseArea())
	Endif

	dbUseArea( .T.,"TOPCONN", TcGenQry( ,,cQuery),cTop02, .T., .T. )

	dbSelectArea(cTop02)

	While (cTop02)->(!Eof())

		aMata250:={;
		{"D3_TM"		,cTM				,NIL},;
		{"D3_COD"		,(cTop02)->PRODUTO	,NIL},;
		{"D3_UM"		,(cTop02)->UM		,NIL},;
		{"D3_QUANT"		,(cTop02)->SALDO	,NIL},;
		{"D3_OP"		,(cTop02)->OP		,NIL},;
		{"D3_LOCAL"		,(cTop02)->ARMAZEM	,NIL},;
		{"D3_OBS"		,"AUTOPPA"			,NIL},;
		{"D3_EMISSAO"	,dDataBase			,NIL}}

		MSExecAuto({|x,y| mata250(x,y)},aMata250,3)

		If lMsErroAuto
			//ConOut("AutOpPa3 - Emp: "+aEmpPrc[1]+" - Filial: "+aEmpPrc[2]+" - Erro para apontar OP: "+SubStr((cTop01)->OP,1,6)+" Produto:"+ AllTrim((cTop01)->PRODUTO) +" - "+dtoc(date())+" - "+time())
			msgStop(MostraErro())
			DisarmTransaction()
		Else
			//ConOut("AutOpPa3 - Emp: "+aEmpPrc[1]+" - Filial: "+aEmpPrc[2]+" - OP: "+SubStr((cTop01)->OP,1,6)+" apontada Produto:"+AllTrim((cTop01)->PRODUTO)+" - Dt.Apont "+dtoc(dDataBase)+" - "+time())
			lMsErroAuto	:=	.F.
		EndIf

		(cTop02)->(dbSkip())
	EndDo

	dDataBase:=	dBkDtBase
	RestArea(aArea04)

Return()

/*=========================================================================================	||
||FUNCAO PARA GERAR NUMERO SEQUENCIAL ORDEM DE PRODUÇÃO										||
||=========================================================================================	*/ 

Static Function NumOpPa()

	Local cNumOp
	Local cMvPar:=	"NUMAUTOP"  

	If !SX6->(dbSeek( Space(TamSx3("C2_FILIAL")[1]) + cMvPar))

		RecLock("SX6", .T.)
		SX6->X6_VAR		:= cMvPar
		SX6->X6_TIPO	:= "C"
		SX6->X6_DESCRIC	:= SX6->X6_DSCSPA  := SX6->X6_DSCENG :="Sequencial OP gerada automaticamente" 
		SX6->X6_DESC1	:= SX6->X6_DSCSPA1 := SX6->X6_DSCENG1:="Sequencial OP gerada automaticamente"
		SX6->X6_DESC2	:= SX6->X6_DSCSPA2 := SX6->X6_DSCENG2:="Sequencial OP gerada automaticamente"
		SX6->X6_PROPRI  := "U"
		cNumOp := SX6->X6_CONTEUD := SX6->X6_CONTSPA := SX6->X6_CONTENG:= "P00001"
		SX6->(MsUnlock())

	Else

		cNumOp := SubStr(SX6->X6_CONTEUD,1,6)

		RecLock("SX6", .F.)		
		cNumOp := Soma1(cNumOp)
		SX6->X6_CONTEUD := cNumOp
		SX6->(MsUnlock())

	EndIf

Return(cNumOp)

/*RELATÓRIO*/
Static Function PRCEXCEL()

	Private cPath	:=	""
	Private cSheet1	:=	""
	Private cTable	:=	"Resumo movimentação"	
	Private oExcel	:=	FWMSEXCEL():New()


	If !ExistDir("C:\TEMP")
		MakeDir("C:\TEMP")
	EndIf

	cPath:= cGetFile("Arquivos xls  (*.xls)  | *.xls  "," ",1,"C:\TEMP",.T.,GETF_LOCALHARD+GETF_RETDIRECTORY ,.F.,.T.)

	If Empty(cPath)
		msgStop('Diretório incorreto!!!','Erro')
		Return()
	Else

		cFilProc:=""

		For nX:=1 to Len(aItens01)

			If aItens01[nX,02] != cFilProc

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

			oExcel:AddRow(cSheet1,cTable,{;
			aItens01[nX,02],;
			aItens01[nX,03],;
			aItens01[nX,04],;
			aItens01[nX,05],;
			aItens01[nX,06],;
			aItens01[nX,07],;
			aItens01[nX,08],;
			aItens01[nX,09],;
			aItens01[nX,10],;
			aItens01[nX,11],;
			aItens01[nX,12],;
			aItens01[nX,13];
			})

		Next

		cArq:= cPath+"RESMOV.xls"

		oExcel:Activate()
		oExcel:GetXMLFile(cArq)
		oExcelApp := MsExcel():New()
		oExcelApp:WorkBooks:Open(cArq)
		oExcelApp:SetVisible(.T.)

	EndIf

Return()
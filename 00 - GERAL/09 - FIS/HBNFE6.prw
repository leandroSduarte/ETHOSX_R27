#Include 'Protheus.ch'
#Include "TOPCONN.CH" 
#Include "Tbiconn.ch"

//-------------------------------------------------------------------
/*/{Protheus.doc} HBNFE6()
@author  Ethosx MOA
@since   16/07/2020
@version 1.0
@obs	Calculo e atualização de campos específicos - NFe
/*/

User Function HBNFE6( )

	Local	aArea      		:= GetArea()
	Local	aPergs			:= { }
	Local	aRet    			:= { }

	Local 	cFilCF			:= cFilAnt
	Local	cMsgDif			:= ""
	Local	cMsgFim		:= ""

	Local 	dDtINi			:= FirstDate(Date())
	Local 	dDtFim			:= LastDate(Date())

	Local	nAliFora			:= 0
	Local	nAliDen			:= 0
	Local	nTipoRegua 	:= 0

	Private		cDifCfo		:= SuperGetMV("MV_XDIFCFO",,"'2406','2407','2551','2556'")

	Private   nTotSFT		:= 0
	Private	   nTotSF3 	:= 0

	aAdd(aPergs, {1, "Selecione a Filial"					,  cFilCF		,  ""			, "!Empty(mv_par01)"	, "SM0"	, ".T."	, 60, 	.F.})
	aAdd(aPergs, {1, "Data De"								,  dDtIni		,  ""			, ".T."							, ""		, ".T."	, 50, 	.F.})
	aAdd(aPergs, {1, "Data Até"							,  dDtFim	,  ""			, ".T."							, ""		, ".T."	, 50, 	.F.})
	aAdd(aPergs, {1, "Aliquota Interestadual"		,  nAliFora	, "@E 99"	, "Positivo()"				, ""		, ".T."	, 50, 	.F.})
	aAdd(aPergs, {1, "Aliquota Interna"				,  nAliDen	, "@E 99"	, "Positivo()"				, ""		, ".T."	, 50, 	.F.})

	If !ParamBox(aPergs, "Calculo do diferencial de alíquota.", @aRet)
		MsgStop("Atenção!"+CHR(10)+CHR(13);
		+"Rotina cancelada pelo usuario!", "HBNFE6")
	Else

		If MsgNoYes("Confirma execução do calculo no período de " + DtoC(aRet[02]) + " até " + DtoC(aRet[03]) + ", na filial " + Trim(aRet[01]) + ;
		"<br><br>[HBNFE6] ", "Atenção...")

			Processa({|| HBNFE6A(aRet)}, "Processando Itens...")
			Processa({|| HBNFE6B(aRet)}, "Processando Cabeçalho...")

			cMsgFim := 'Processamento Finalizado na Filial ' + Trim(aRet[01]) + '.'
			cMsgFim += '<br>'
			cMsgFim += 'De ' + DtoC(aRet[02]) + ' até ' + DtoC(aRet[03]) + '.'
			cMsgFim += '<br>'
			cMsgFim += '<br>'
			cMsgFim += 'Total de Produtos atualizados (SFT): <b><font color="#FF0000">' + str(nTotSFT) + '</font></b>'
			cMsgFim += '<br>'
			cMsgFim += 'Total de Documentos Atualizados (SF3): <b><font color="#FF0000">' + str(nTotSF3) + '</font></b>'
			cMsgFim += '<br>'
			cMsgFim += '<br>'
			cMsgFim += '[HBNFE6]'

			MsgAlert(cMsgFim, "Atenção!....")

		EndIf

	EndIf

	RestArea(aArea)

Return Nil

/*-----------------------------------------------------------*
| Func.: HBNFE6A                                                   |
| Desc.: Processamento do calculo diferencial           |
*-----------------------------------------------------------*/
Static Function HBNFE6A(aRet)

	Local	aArea  	:= GetArea()

	Local	cAlias	:= ""
	Local	cQry		:= ""
	Local	cQrySFT:= ""
	Local	cUf		:= ""

	Local 	dDti		
	Local 	dDtf	

	Local	nAtual 		:= 0
	Local	nTotal 		:= 0
	Local	nValMerc 	:= 0
	Local	nValTot		:= 0
	Local	nValIcm 	:= 0
	Local	nDifal		:= 0
	Local	nIcmsDif	:= 0	
	Local	nExec		:= 0

	dDti 	:= DtoS(aRet[02])
	dDtf 	:= DtoS(aRet[03])

	If Select("QRY_AUX") > 0
		QRY_AUX->(DBCloseArea(  ))
	EndIf

	cQry	:= " 	SELECT "
	cQry	+= " 		FT_FILIAL, FT_NFISCAL, FT_SERIE, FT_ENTRADA, FT_CLIEFOR, FT_LOJA, FT_ESTADO, 
	cQry	+= " 		FT_CFOP, FT_VALCONT, FT_VALICM, FT_ICMSCOM, FT_BASEICM, FT_PRODUTO, FT_ITEM, B1_DESC,X5_DESCENG"
	cQry	+= " FROM " + RETSqlName("SFT") + " SFT "
	cQry	+= " LEFT JOIN " + RETSqlName("SB1") + " SB1 ON B1_COD = FT_PRODUTO AND SB1.D_E_L_E_T_ = '' "
	cQry	+= " LEFT JOIN " + RETSqlName("SX5") + " SX5 ON X5_TABELA = '_F' AND X5_FILIAL = FT_FILIAL AND SX5.D_E_L_E_T_ = '' "  
	cQry	+= " WHERE "
	cQry	+= "		FT_FILIAL = '" + Trim(aRet[01]) + "' AND " 
	cQry	+= "   	FT_ENTRADA BETWEEN '" + dDti + "'  AND '" + dDtf + "' AND  "
	cQry	+= "   	FT_CFOP IN (" + Trim(cDifCfo) + ") AND  "
	cQry	+= "   	SFT.D_E_L_E_T_ = '' "
	// memowrit("c:\siga\cqrydifal.txt",cQry)	

	//Executa a consulta
	TCQuery cQry New Alias "QRY_AUX"

	//Conta quantos registros existem, e seta no tamanho da régua
	Count To nTotal
	ProcRegua(nTotal)

	//Percorre todos os registros da query
	QRY_AUX->(DbGoTop())
	While ! QRY_AUX->(EoF())

		//Incrementa a mensagem na régua
		nAtual++
		IncProc("Analisando registro " + cValToChar(nAtual) + " de " + cValToChar(nTotal) + "...")

		cUf := Trim(QRY_AUX->X5_DESCENG)

		nValMerc 	:= QRY_AUX->FT_VALCONT
		//nValIcm 	:= QRY_AUX->FT_VALICMS -- VERIFICAR COM ISABEL - VALOR ZERO!!!
		nValIcm 	:= ( (QRY_AUX->FT_VALCONT * aRet[04]) / 100 )
		nValTot 	:= ( QRY_AUX->FT_VALCONT + nValIcm )

		If  cUf $ ("BA/MG/PA/PR/RS")

			nDifal := nValMerc
			nDifal := nDifal * ( ( 100 - aRet[04] ) / 100 )
			nDifal := nDifal / ( ( 100 - aRet[05] ) / 100 )
			nDifal := Ndifal - nValMerc

			cQrySFT:= ""
			cQrySFT:= " 	UPDATE " + RETSqlName("SFT")
			cQrySFT+= " 		SET FT_ICMSCOM = ROUND(" + str(nDifal) + ",2)"
			cQrySFT+= "FROM " + RETSqlName("SFT") + " SFT "
			cQrySFT+= "WHERE "
			cQrySFT+= "		FT_FILIAL = '" + QRY_AUX->FT_FILIAL + "' AND "
			cQrySFT+= "		FT_NFISCAL = '" + QRY_AUX->FT_NFISCAL + "' AND  "
			cQrySFT+= "		FT_SERIE = '" + QRY_AUX->FT_SERIE + "' AND  "
			cQrySFT+= " 	FT_ENTRADA = '" + QRY_AUX->FT_ENTRADA + "' AND " 
			cQrySFT+= "		FT_CLIEFOR = '" + QRY_AUX->FT_CLIEFOR + "' AND "
			cQrySFT+= "		FT_LOJA = '" + QRY_AUX->FT_LOJA + "' AND "
			cQrySFT+= "		FT_ESTADO = '" + QRY_AUX->FT_ESTADO + "' AND "
			cQrySFT+= "		FT_CFOP = '" + QRY_AUX->FT_CFOP + "' AND  "
			cQrySFT+= "		FT_PRODUTO = '" + QRY_AUX->FT_PRODUTO + "' AND "
			cQrySFT+= "		FT_ITEM = '" + QRY_AUX->FT_ITEM + "' AND "
			cQrySFT+= "		SFT.D_E_L_E_T_ = ''  "
			// memowrit("c:\siga\cqrydifal_upd_sft_01.txt",cQrySFT)	

			Begin Transaction
				nExec	:= TCSqlExec(cQrySFT)

				If (nExec < 0)
					DisarmTransaction()
					MsgStop("TCSQLError() " + TCSQLError(), "Erro na atualização do Difal - Itens!")
				Else
					IncProc("Atualizando registro " + cValToChar(nAtual) + " de " + cValToChar(nTotal) + "...")
				EndIf

			End Transaction


			/*
			cMsgDif := 'Aliq. Interestadual:   <b>' + Str(aRet[04]) + '</b>  -  Aliq. Interna:   <b>' + Str(aRet[05]) + '</b>.'
			cMsgDif += '<br>'
			cMsgDif += 'Filial:   ' + Trim(aRet[01]) + '.'
			cMsgDif += '<br>'
			cMsgDif += 'Nota Fiscal / Série:   ' + QRY_AUX->FT_NFISCAL+ ' / ' + QRY_AUX->FT_SERIE + '.'
			cMsgDif += '<br>'
			cMsgDif += 'Cliente / Loja / Estado:   ' + QRY_AUX->FT_CLIEFOR + ' / ' + QRY_AUX->FT_LOJA + ' / ' + QRY_AUX->X5_DESCENG
			cMsgDif += '<br>'
			cMsgDif += 'CFOP:   ' + QRY_AUX->FT_CFOP + '.'
			cMsgDif += '<br>'
			cMsgDif += 'Produto:   ' + QRY_AUX->FT_PRODUTO + ' - ' + QRY_AUX->B1_DESC
			cMsgDif += '<br>'
			cMsgDif += 'Valor Total Produto:   ' + Transform(QRY_AUX->FT_VALCONT	,"@E 999,999.99")+ '.'   
			cMsgDif += '<br>'
			cMsgDif += 'Valor Icms:   ' + Transform(QRY_AUX->FT_VALICM	,"@E 999,999.99") + '.'
			cMsgDif += '<br>'
			cMsgDif += 'Valor FT_ICMSCOM atual:   <b><font color="#FF0000">' + Transform(QRY_AUX->FT_ICMSCOM  ,"@E 999,999.99")+ '.</font></b>'
			cMsgDif += '<br>'
			cMsgDif += 'Valor calculado:   <b><font color="#FF0000">' + Transform(nDifal  ,"@E 999,999.99") + '.</font></b>'

			MsgAlert(cMsgDif, "Comparativo!....")
			*/

		ElseIf cUf = "GO"

			nIcmsDif	:= ( aRet[05] - aRet[04])
			nDifal 		:= nValTot 
			nDifal 		:= ( nDifal / ( (100 - aRet[05]) / 100) )
			nDifal 		:= ( (nDifal * nIcmsDif) /100 )

			cQrySFT:= ""
			cQrySFT:= " 	UPDATE " + RETSqlName("SFT")
			cQrySFT+= " 		SET FT_ICMSCOM = ROUND(" + str(nDifal) + ",2)"
			cQrySFT+= "FROM " + RETSqlName("SFT") + " SFT "
			cQrySFT+= "WHERE "
			cQrySFT+= "		FT_FILIAL = '" + QRY_AUX->FT_FILIAL + "' AND "
			cQrySFT+= "		FT_NFISCAL = '" + QRY_AUX->FT_NFISCAL + "' AND  "
			cQrySFT+= "		FT_SERIE = '" + QRY_AUX->FT_SERIE + "' AND  "
			cQrySFT+= " 	FT_ENTRADA = '" + QRY_AUX->FT_ENTRADA + "' AND " 
			cQrySFT+= "		FT_CLIEFOR = '" + QRY_AUX->FT_CLIEFOR + "' AND "
			cQrySFT+= "		FT_LOJA = '" + QRY_AUX->FT_LOJA + "' AND "
			cQrySFT+= "		FT_ESTADO = '" + QRY_AUX->FT_ESTADO + "' AND "
			cQrySFT+= "		FT_CFOP = '" + QRY_AUX->FT_CFOP + "' AND  "
			cQrySFT+= "		FT_PRODUTO = '" + QRY_AUX->FT_PRODUTO + "' AND "
			cQrySFT+= "		FT_ITEM = '" + QRY_AUX->FT_ITEM + "' AND "
			cQrySFT+= "		SFT.D_E_L_E_T_ = ''  "
			// memowrit("c:\siga\cqrydifal_upd_sft_02.txt",cQrySFT)	

			Begin Transaction
				nExec	:= TCSqlExec(cQrySFT)

				If (nExec < 0)
					DisarmTransaction()
					MsgStop("TCSQLError() " + TCSQLError(), "Erro na atualização do Difal - Itens!")
				Else
					IncProc("Atualizando registro " + cValToChar(nAtual) + " de " + cValToChar(nTotal) + "...")
				EndIf

			End Transaction

			/*
			cMsgDif := 'Aliq. Interestadual:   <b>' + Str(aRet[04]) + '</b>  -  Aliq. Interna:   <b>' + Str(aRet[05]) + '</b>.'
			cMsgDif += '<br>'
			cMsgDif += 'Filial:   ' + Trim(aRet[01]) + '.'
			cMsgDif += '<br>'
			cMsgDif += 'Nota Fiscal / Série:   ' + QRY_AUX->FT_NFISCAL+ ' / ' + QRY_AUX->FT_SERIE + '.'
			cMsgDif += '<br>'
			cMsgDif += 'Cliente / Loja / Estado:   ' + QRY_AUX->FT_CLIEFOR + ' / ' + QRY_AUX->FT_LOJA + ' / ' + QRY_AUX->X5_DESCENG
			cMsgDif += '<br>'
			cMsgDif += 'CFOP:   ' + QRY_AUX->FT_CFOP + '.'
			cMsgDif += '<br>'
			cMsgDif += 'Produto:   ' + QRY_AUX->FT_PRODUTO + ' - ' + QRY_AUX->B1_DESC
			cMsgDif += '<br>'
			cMsgDif += 'Valor Total Produto:   ' + Transform(QRY_AUX->FT_VALCONT	,"@E 999,999.99")+ '.'   
			cMsgDif += '<br>'
			cMsgDif += 'Valor Icms:   ' + Transform(QRY_AUX->FT_VALICM	,"@E 999,999.99") + '.'
			cMsgDif += '<br>'
			cMsgDif += 'Valor FT_ICMSCOM atual:   <b><font color="#FF0000">' + Transform(QRY_AUX->FT_ICMSCOM  ,"@E 999,999.99")+ '.</font></b>'
			cMsgDif += '<br>'
			cMsgDif += 'Valor calculado:   <b><font color="#FF0000">' + Transform(nDifal  ,"@E 999,999.99") + '.</font></b>'

			MsgAlert(cMsgDif, "Comparativo!....")
			*/
		Else

			nDifal := ( (aRet[05] - aRet[04]) * QRY_AUX->FT_VALCONT ) / 100

			cQrySFT:= ""
			cQrySFT:= " 	UPDATE " + RETSqlName("SFT")
			cQrySFT+= " 		SET FT_ICMSCOM = ROUND(" + str(nDifal) + ",2)"
			cQrySFT+= "FROM " + RETSqlName("SFT") + " SFT "
			cQrySFT+= "WHERE "
			cQrySFT+= "		FT_FILIAL = '" + QRY_AUX->FT_FILIAL + "' AND "
			cQrySFT+= "		FT_NFISCAL = '" + QRY_AUX->FT_NFISCAL + "' AND  "
			cQrySFT+= "		FT_SERIE = '" + QRY_AUX->FT_SERIE + "' AND  "
			cQrySFT+= " 	FT_ENTRADA = '" + QRY_AUX->FT_ENTRADA + "' AND " 
			cQrySFT+= "		FT_CLIEFOR = '" + QRY_AUX->FT_CLIEFOR + "' AND "
			cQrySFT+= "		FT_LOJA = '" + QRY_AUX->FT_LOJA + "' AND "
			cQrySFT+= "		FT_ESTADO = '" + QRY_AUX->FT_ESTADO + "' AND "
			cQrySFT+= "		FT_CFOP = '" + QRY_AUX->FT_CFOP + "' AND  "
			cQrySFT+= "		FT_PRODUTO = '" + QRY_AUX->FT_PRODUTO + "' AND "
			cQrySFT+= "		FT_ITEM = '" + QRY_AUX->FT_ITEM + "' AND "
			cQrySFT+= "		SFT.D_E_L_E_T_ = ''  "
			// memowrit("c:\siga\cqrydifal_upd_sft_03.txt",cQrySFT)	

			Begin Transaction
				nExec	:= TCSqlExec(cQrySFT)

				If (nExec < 0)
					DisarmTransaction()
					MsgStop("TCSQLError() " + TCSQLError(), "Erro na atualização do Difal - Itens!")
				Else
					IncProc("Atualizando registro " + cValToChar(nAtual) + " de " + cValToChar(nTotal) + "...")
				EndIf

			End Transaction

			/*
			cMsgDif := 'Aliq. Interestadual:   <b>' + Str(aRet[04]) + '</b>  -  Aliq. Interna:   <b>' + Str(aRet[05]) + '</b>.'
			cMsgDif += '<br>'
			cMsgDif += 'Filial:   ' + Trim(aRet[01]) + '.'
			cMsgDif += '<br>'
			cMsgDif += 'Nota Fiscal / Série:   ' + QRY_AUX->FT_NFISCAL+ ' / ' + QRY_AUX->FT_SERIE + '.'
			cMsgDif += '<br>'
			cMsgDif += 'Cliente / Loja / Estado:   ' + QRY_AUX->FT_CLIEFOR + ' / ' + QRY_AUX->FT_LOJA + ' / ' + QRY_AUX->X5_DESCENG
			cMsgDif += '<br>'
			cMsgDif += 'CFOP:   ' + QRY_AUX->FT_CFOP + '.'
			cMsgDif += '<br>'
			cMsgDif += 'Produto:   ' + QRY_AUX->FT_PRODUTO + ' - ' + QRY_AUX->B1_DESC
			cMsgDif += '<br>'
			cMsgDif += 'Valor Total Produto:   ' + Transform(QRY_AUX->FT_VALCONT	,"@E 999,999.99")+ '.'   
			cMsgDif += '<br>'
			cMsgDif += 'Valor Icms:   ' + Transform(QRY_AUX->FT_VALICM	,"@E 999,999.99") + '.'
			cMsgDif += '<br>'
			cMsgDif += 'Valor FT_ICMSCOM atual:   <b><font color="#FF0000">' + Transform(QRY_AUX->FT_ICMSCOM  ,"@E 999,999.99")+ '.</font></b>'
			cMsgDif += '<br>'
			cMsgDif += 'Valor calculado:   <b><font color="#FF0000">' + Transform(nDifal  ,"@E 999,999.99") + '.</font></b>'

			MsgAlert(cMsgDif, "Comparativo!....")
			*/
		EndIf	

		QRY_AUX->(DbSkip())

	EndDo
	
	nTotSFT = nAtual

	QRY_AUX->(DbCloseArea())

	RestArea(aArea)

Return

/*-----------------------------------------------------------*
| Func.: HBNFE6B                                                   |
| Desc.: Processamento do calculo diferencial           |
*-----------------------------------------------------------*/
Static Function HBNFE6B(aRet)

	Local	aArea  	:= GetArea()

	Local	cAlias	:= ""
	Local	cQry		:= ""
	Local	cQrySF3:= ""
	Local	cUf		:= ""

	Local 	dDti		
	Local 	dDtf	

	Local	nAtual 		:= 0
	Local	nTotal 		:= 0
	Local 	nOldDif		:= 0
	Local	nExec		:= 0

	dDti 	:= DtoS(aRet[02])
	dDtf 	:= DtoS(aRet[03])

	If Select("QRY_AUX") > 0
		QRY_AUX->(DBCloseArea(  ))
	EndIf

	cQry	:= " 	SELECT "
	cQry	+= " 		F3_FILIAL, F3_NFISCAL, F3_SERIE, F3_ENTRADA, F3_CLIEFOR, F3_LOJA, F3_ESTADO, F3_CFO, F3_VALCONT, F3_ICMSCOM"
	cQry	+= " FROM " + RETSqlName("SF3") + " SF3 "
	cQry	+= " WHERE "
	cQry	+= "		F3_FILIAL = '" + Trim(aRet[01]) + "' AND " 
	cQry	+= "   	F3_ENTRADA BETWEEN '" + dDti + "'  AND '" + dDtf + "' AND  "
	cQry	+= "   	F3_CFO IN (" + Trim(cDifCfo) + ") AND  "
	cQry	+= "   	SF3.D_E_L_E_T_ = '' "
	memowrit("c:\siga\cqrydifal_SF3.txt",cQry)	

	//Executa a consulta	TCQuery cQry New Alias "QRY_AUX"
	TCQuery cQry New Alias "QRY_AUX"

	//Conta quantos registros existem, e seta no tamanho da régua
	Count To nTotal
	ProcRegua(nTotal)

	//Percorre todos os registros da query
	QRY_AUX->(DbGoTop())
	While ! QRY_AUX->(EoF())

		//Incrementa a mensagem na régua
		nAtual++
		IncProc("Analisando registro " + cValToChar(nAtual) + " de " + cValToChar(nTotal) + "...")

		//Gravando valor atual do Diferencia
		nOldDif := QRY_AUX->F3_ICMSCOM

		cQrySF3:= ""
		cQrySF3:= "	UPDATE " + RETSqlName("SF3")
		cQrySF3+= " 		SET F3_ICMSCOM = "
		cQrySF3+= " 		ROUND( (SELECT SUM(FT_ICMSCOM) FROM " + RETSqlName("SFT") + " SFT WHERE FT_FILIAL = F3_FILIAL AND FT_NFISCAL = F3_NFISCAL AND  "
		cQrySF3+= " 		FT_SERIE = F3_SERIE AND FT_ENTRADA = F3_ENTRADA AND FT_CLIEFOR = F3_CLIEFOR AND FT_LOJA = F3_LOJA AND SFT.D_E_L_E_T_ = '' "
		cQrySF3+= " 		GROUP BY  "
		cQrySF3+= " 		FT_FILIAL, FT_NFISCAL, FT_SERIE, FT_ENTRADA, FT_CLIEFOR, FT_LOJA),2 ) "
		cQrySF3+= " 		FROM  " + RETSqlName("SF3") + " SF3 "
		cQrySF3+= " 		WHERE  "
		cQrySF3+= " 		F3_FILIAL = '" + QRY_AUX->F3_FILIAL + "' AND "
		cQrySF3+= " 		F3_NFISCAL = '" + QRY_AUX->F3_NFISCAL + "' AND "
		cQrySF3+= " 		F3_SERIE = '" + QRY_AUX->F3_SERIE + "' AND "
		cQrySF3+= " 		F3_ENTRADA = '" + QRY_AUX->F3_ENTRADA + "' AND "
		cQrySF3+= " 		F3_CLIEFOR = '" + QRY_AUX->F3_CLIEFOR + "' AND "
		cQrySF3+= " 		F3_LOJA = '" + QRY_AUX->F3_LOJA + "' AND "
		cQrySF3+= " 		F3_ESTADO = '" + QRY_AUX->F3_ESTADO + "' AND "
		cQrySF3+= " 		F3_CFO = '" + QRY_AUX->F3_CFO + "' AND "
		cQrySF3+= " 		SF3.D_E_L_E_T_ = ''  "
		memowrit("c:\siga\cqrydifal_upd_sf3.txt",cQrySF3)	

		Begin Transaction
			nExec	:= TCSqlExec(cQrySF3)

			If (nExec < 0)
				DisarmTransaction()
				MsgStop("TCSQLError() " + TCSQLError(), "Erro na atualização do Difal - Itens!")
			Else
				IncProc("Atualizando registro " + cValToChar(nAtual) + " de " + cValToChar(nTotal) + "...")
			EndIf

		End Transaction

		/*
		cMsgDif := 'Aliq. Interestadual:   <b>' + Str(aRet[04]) + '</b>  -  Aliq. Interna:   <b>' + Str(aRet[05]) + '</b>.'
		cMsgDif += '<br>'
		cMsgDif += 'Filial:   ' + Trim(aRet[01]) + '.'
		cMsgDif += '<br>'
		cMsgDif += 'Nota Fiscal / Série:   ' + QRY_AUX->F3_NFISCAL+ ' / ' + QRY_AUX->F3_SERIE + '.'
		cMsgDif += '<br>'
		cMsgDif += 'Cliente / Loja / Estado:   ' + QRY_AUX->F3_CLIEFOR + ' / ' + QRY_AUX->F3_LOJA 
		cMsgDif += '<br>'
		cMsgDif += 'CFOP:   ' + QRY_AUX->F3_CFO + '.'
		cMsgDif += '<br>'
		cMsgDif += 'Valor Total Documento:   ' + Transform(QRY_AUX->F3_VALCONT	,"@E 999,999.99")+ '.'   
		cMsgDif += '<br>'
		cMsgDif += 'Valor F3_ICMSCOM atual:   <b><font color="#FF0000">' + Transform(nOldDif ,"@E 999,999.99")+ '.</font></b>'
		cMsgDif += '<br>'
		cMsgDif += 'Valor calculado:   <b><font color="#FF0000">' + Transform(QRY_AUX->F3_ICMSCOM   ,"@E 999,999.99") + '.</font></b>'

		MsgAlert(cMsgDif, "Comparativo!....")
		*/

		QRY_AUX->(DbSkip())

	EndDo
	
	nTotSF3 = nAtual
	
	QRY_AUX->(DbCloseArea())

	RestArea(aArea)

Return
#Include 'Protheus.ch'

//00=Nao Utilizado;01=HE dia util;02=HE dia nao util;03=Banco de Horas;04=Feriado;05=DSR Total                                    
/*------------------------------------------------------------------------------
|Programa: 		XRELHE															| 
|Tipo: 			Combo campo RV_XRELHE                      						|
|Empresa: 		Habib's 														|
|Analista: 		Bruno															|
|Consultoria:	Ethosx															|
-------------------------------------------------------------------------------*/
 
User Function crvrelhe()
Local aArea   := GetArea()
Local cOpcoes := ""
                                                                        
//Montando as opções de retorno
cOpcoes += "00=Não Utilizado;"
cOpcoes += "01=HE dia util;"
cOpcoes += "02=HE dia nao util;"
cOpcoes += "03=Banco de Horas;"
cOpcoes += "04=Feriado;"
cOpcoes += "05=DSR Total;"
 
RestArea(aArea)
Return cOpcoes



/*------------------------------------------------------------------------------
|Programa: 		XRELHE															| 
|Tipo: 			Relatorio de Custos Horas Extras								|
|Empresa: 		Habib's 														|
|Analista: 		Bruno															|
|Consultoria:	Ethosx															|
-------------------------------------------------------------------------------*/

User Function XRELHE()

Private cPerg	   := PadR("XRELHEC",10)
Private cNextAlias := GetNextAlias()
Private cTipo 	   := ()

ValidPerg(cPerg)

If Pergunte(cPerg , .T.)

	oReport:= ReportDef()
	oReport:PrintDialog()

EndIf

Return

/*------------------------------------------------------------------------------
|Programa: 		XRELHE															| 
|Tipo: 			Relatorio de Custos Horas Extras								|
|Empresa: 		Habib's 														|
|Analista: 		Bruno															|
|Consultoria:	Ethosx															|
-------------------------------------------------------------------------------*/

Static Function ReportDef()      

oReport := TReport():New(cPerg,"Relatorio Horas Extras",cPerg, {|oReport| ReportPrint(oReport)},"Custos de Horas Extras")
oReport:SetLandscape(.T.)
oReport:SetPortrait()
oReport:HideParamPage()
oReport:HideHeader()
oReport:HideFooter()
oReport:SetDevice(4) 	  // Planilha Excel
oReport:SetEnvironment(2) // Local
oReport:cFontBody := "Calibri"
oReport:nFontBody := 9
oReport:SetFile("")

oSection1:= TRSection():New(oReport,OemToAnsi("Custo de Horas Extras"), {"SRD","SRA","SRV","CTE"})

//Dados Gerais
TRCell():New(oSection1,"FILIAL" 	,"CTE",	"Filial"      				,/*Picture*/,/*Tamanho*/,/*lPixel*/,{|| (cNextAlias)->FILIAL    				})
TRCell():New(oSection1,"DESC_FILIAL","CTE",	"Nome Filial"      			,/*Picture*/,/*Tamanho*/,/*lPixel*/,{|| FwFilName(cEmpAnt, (cNextAlias)->FILIAL)})
TRCell():New(oSection1,"MATRICULA"  ,"CTE",	"Matricula"   				,/*Picture*/,/*Tamanho*/,/*lPixel*/,{|| (cNextAlias)->MATRICULA 				})
TRCell():New(oSection1,"NOME"  		,"CTE",	"Nome Colaborador"   		,/*Picture*/,/*Tamanho*/,/*lPixel*/,{|| (cNextAlias)->NOME 						})
TRCell():New(oSection1,"CPF"  		,"CTE",	"CPF"   					,/*Picture*/,/*Tamanho*/,/*lPixel*/,{|| (cNextAlias)->CPF 						})
TRCell():New(oSection1,"CC"  		,"CTE",	"Cod. Centro Custo"   		,/*Picture*/,/*Tamanho*/,/*lPixel*/,{|| (cNextAlias)->CC 						})
TRCell():New(oSection1,"DESCCC"  	,"CTE",	"Denominacao Centro Custo"  ,/*Picture*/,/*Tamanho*/,/*lPixel*/,{|| (cNextAlias)->DESCCC 					})
TRCell():New(oSection1,"CC"  		,"CTE",	"Cod.Lotacao"   			,/*Picture*/,/*Tamanho*/,/*lPixel*/,{|| (cNextAlias)->CC 						})
TRCell():New(oSection1,"DESCCC"  	,"CTE",	"Denominacao Lotacao"  		,/*Picture*/,/*Tamanho*/,/*lPixel*/,{|| (cNextAlias)->DESCCC 					})
TRCell():New(oSection1,"SITUACAO"  	,"CTE",	"Situacao"   				,/*Picture*/,/*Tamanho*/,/*lPixel*/,{|| (cNextAlias)->SITUACAO 					})
TRCell():New(oSection1,"FUNCAO"  	,"CTE",	"Denominacao Cargo"   		,/*Picture*/,/*Tamanho*/,/*lPixel*/,{|| (cNextAlias)->FUNCAO 					})
TRCell():New(oSection1,"PESSOA"  	,"CTE",	"Denominacao Tipo Pessoa"   ,/*Picture*/,/*Tamanho*/,/*lPixel*/,{|| (cNextAlias)->PESSOA 					})
TRCell():New(oSection1,"HEH" 		,"CTE1","Ref - HE dia util"       	,/*Picture*/,/*Tamanho*/,/*lPixel*/,{|| (cNextAlias)->HEH 	 					})
TRCell():New(oSection1,"HE" 		,"CTE",	"Val - HE dia util"       	,/*Picture*/,/*Tamanho*/,/*lPixel*/,{|| (cNextAlias)->HE 	 					})
TRCell():New(oSection1,"HENH" 		,"CTE1","Ref - HE dia nao util"   	,/*Picture*/,/*Tamanho*/,/*lPixel*/,{|| (cNextAlias)->HENH 	 					})
TRCell():New(oSection1,"HEN" 		,"CTE",	"Val - HE dia nao util"   	,/*Picture*/,/*Tamanho*/,/*lPixel*/,{|| (cNextAlias)->HEN 	 					})
TRCell():New(oSection1,"BDH" 		,"CTE1","Ref - Banco de Horas"    	,/*Picture*/,/*Tamanho*/,/*lPixel*/,{|| (cNextAlias)->BDH	 					})
TRCell():New(oSection1,"BH" 		,"CTE",	"Val - Banco de Horas"    	,/*Picture*/,/*Tamanho*/,/*lPixel*/,{|| (cNextAlias)->BH	 					})
TRCell():New(oSection1,"FERIADOH" 	,"CTE1","Ref - Feriado" 			,/*Picture*/,/*Tamanho*/,/*lPixel*/,{|| (cNextAlias)->FERIADOH  				})
TRCell():New(oSection1,"FERIADO" 	,"CTE",	"Val - Feriado" 			,/*Picture*/,/*Tamanho*/,/*lPixel*/,{|| (cNextAlias)->FERIADO   				})
TRCell():New(oSection1,"DSR"   		,"CTE",	"DSR Total"         		,/*Picture*/,/*Tamanho*/,/*lPixel*/,{|| (cNextAlias)->DSR  	 					})
TRCell():New(oSection1,"TOTALH"		,"CTE",	"Total Horas"      			,/*Picture*/,/*Tamanho*/,/*lPixel*/,{|| (cNextAlias)->TOTALH 					})
TRCell():New(oSection1,"TOTALV"		,"CTE",	"Total Valores"				,/*Picture*/,/*Tamanho*/,/*lPixel*/,{|| (cNextAlias)->TOTALV					})

Return oReport

/*------------------------------------------------------------------------------
|Programa: 		XRELHE															| 
|Tipo: 			Relatorio de Custos Horas Extras								|
|Empresa: 		Habib's 														|
|Analista: 		Bruno															|
|Consultoria:	Ethosx															|
-------------------------------------------------------------------------------*/

Static Function ReportPrint(oReport)

Local oSection	:= oReport:Section(1)
Local cQuery	:= ""



cQuery += " WITH CTE AS (SELECT  " + CRLF
cQuery += " RD_FILIAL AS FILIAL, " + CRLF
cQuery += " RD_MAT AS MATRICULA, RA_NOME AS NOME,RA_CIC AS CPF,RD_CC AS CC,CTT_DESC01 AS DESCCC, " + CRLF
cQuery += " CASE WHEN RA_SITFOLH = '' THEN 'ATIVO' " + CRLF
cQuery += " WHEN RA_SITFOLH = 'D' THEN 'DEMITIDO' " + CRLF
cQuery += " WHEN RA_SITFOLH = 'F' THEN 'FERIAS' " + CRLF
cQuery += " WHEN RA_SITFOLH = 'A' THEN 'AFASTADO' END AS SITUACAO, " + CRLF
cQuery += " RJ_DESC AS FUNCAO, " + CRLF
cQuery += " CASE " + CRLF
cQuery += " WHEN RA_CATFUNC = 'M' THEN 'MENSALISTA' " + CRLF
cQuery += " WHEN RA_CATFUNC = 'P' THEN 'PRO LABORE' " + CRLF
cQuery += " WHEN RA_CATFUNC = 'H' THEN 'HORISTA' " + CRLF
cQuery += " WHEN RA_CATFUNC = 'A' THEN 'AUTONOMO' " + CRLF
cQuery += " WHEN RA_CATFUNC = 'C' THEN 'COMISSIONADO' " + CRLF
cQuery += " WHEN RA_CATFUNC = 'E' THEN 'ESTAGIARIO MENSALISTA' " + CRLF
cQuery += " WHEN RA_CATFUNC = 'G' THEN 'ESTAGIARIO HORISTA' " +  CRLF
cQuery += " ELSE '' " + CRLF
cQuery += " END AS PESSOA, " + CRLF
cQuery += " CASE WHEN [01] IS NULL THEN '0' ELSE [01] END AS HE, " + CRLF
cQuery += " CASE WHEN [02] IS NULL THEN '0' ELSE [02] END AS HEN, " + CRLF
cQuery += " CASE WHEN [03] IS NULL THEN '0' ELSE [03] END AS BH, " + CRLF
cQuery += " CASE WHEN [04] IS NULL THEN '0' ELSE [04] END AS FERIADO, " + CRLF
cQuery += " CASE WHEN [05] IS NULL THEN '0' ELSE [05] END AS DSR " + CRLF
 
cQuery += " FROM (SELECT RD_FILIAL,RD_MAT,RA_NOME,RA_CIC,RD_CC,CTT_DESC01,RA_SITFOLH,RJ_DESC,RA_CATFUNC, RV_XRELHE AS TIPO,SUM(RD_VALOR) AS VALOR  FROM " +  retSqlname("SRD") + " SRD " + CRLF
		    
cQuery += " LEFT JOIN " + retSqlname("SRA") + " SRA " + CRLF
cQuery += " ON RA_FILIAL = RD_FILIAL AND RA_MAT = RD_MAT AND SRA.D_E_L_E_T_ = ''  " + CRLF
		    
cQuery += " LEFT JOIN " + retSqlname("SRV") + " SRV " + CRLF
cQuery += " ON RV_FILIAL = '" + xFilial("SRV") + "' AND RV_COD = RD_PD AND SRV.D_E_L_E_T_ = '' " + CRLF

cQuery += " LEFT JOIN " + retSqlname("CTT") + " CTT " + CRLF
cQuery += " ON CTT_FILIAL = '" + xFilial("CTT") + "' AND CTT_CUSTO = RD_CC AND CTT.D_E_L_E_T_ = '' " + CRLF

cQuery += " LEFT JOIN " + retSqlname("SRJ") + " SRJ " + CRLF
cQuery += " ON RJ_FILIAL = '" + xFilial("SRJ") + "' AND RA_CODFUNC = RJ_FUNCAO AND SRJ.D_E_L_E_T_ = '' " + CRLF
		    
cQuery += " WHERE SRD.D_E_L_E_T_ = '' " + CRLF
cQuery += " AND RD_PERIODO 		  = '" + MV_PAR02 + "' AND RD_ROTEIR = '" + MV_PAR01 + "' AND RD_SEMANA = '" + MV_PAR03 + "' " + CRLF
cQuery += " AND RD_FILIAL 	BETWEEN '" + MV_PAR04 + "' AND '" + MV_PAR05 + "' " + CRLF
cquery += " AND RD_MAT 		BETWEEN '" + MV_PAR06 + "' AND '" + MV_PAR07 + "' " + CRLF
cquery += " AND RD_CC 		BETWEEN '" + MV_PAR08 + "' AND '" + MV_PAR09 + "' " + CRLF
		    
cQuery += " GROUP BY RD_FILIAL,RD_MAT,RV_XRELHE,RA_NOME,RA_CIC,RD_CC,CTT_DESC01,RA_SITFOLH,RJ_DESC,RA_CATFUNC) P  " + CRLF
		    
cQuery += " PIVOT (SUM(VALOR) FOR TIPO IN ([00],[01],[02],[03],[04],[05])) AS PVT " + CRLF

cQuery += " UNION ALL " + CRLF

cQuery += " SELECT  " + CRLF
cQuery += " RC_FILIAL AS FILIAL, " + CRLF  
cQuery += " RC_MAT AS MATRICULA, RA_NOME AS NOME,RA_CIC AS CPF,RC_CC AS CC,CTT_DESC01 AS DESCCC, " + CRLF
cQuery += " CASE WHEN RA_SITFOLH = '' THEN 'ATIVO' " + CRLF
cQuery += " WHEN RA_SITFOLH = 'D' THEN 'DEMITIDO'  " + CRLF
cQuery += " WHEN RA_SITFOLH = 'F' THEN 'FERIAS' " + CRLF
cQuery += " WHEN RA_SITFOLH = 'A' THEN 'AFASTADO' END AS SITUACAO, " + CRLF
cQuery += " RJ_DESC AS FUNCAO, " + CRLF
cQuery += " CASE " + CRLF
cQuery += " WHEN RA_CATFUNC = 'M' THEN 'MENSALISTA'" + CRLF
cQuery += " WHEN RA_CATFUNC = 'P' THEN 'PRO LABORE'" + CRLF
cQuery += " WHEN RA_CATFUNC = 'H' THEN 'HORISTA'" + CRLF
cQuery += " WHEN RA_CATFUNC = 'A' THEN 'AUTONOMO'" + CRLF
cQuery += " WHEN RA_CATFUNC = 'C' THEN 'COMISSIONADO'" + CRLF
cQuery += " WHEN RA_CATFUNC = 'E' THEN 'ESTAGIARIO MENSALISTA'" + CRLF
cQuery += " WHEN RA_CATFUNC = 'G' THEN 'ESTAGIARIO HORISTA'" + CRLF
cQuery += " ELSE '' " + CRLF
cQuery += " END AS PESSOA, " + CRLF
cQuery += " CASE WHEN [01] IS NULL THEN '0' ELSE [01] END AS HE, " + CRLF
cQuery += " CASE WHEN [02] IS NULL THEN '0' ELSE [02] END AS HEN, " + CRLF
cQuery += " CASE WHEN [03] IS NULL THEN '0' ELSE [03] END AS BH, " + CRLF
cQuery += " CASE WHEN [04] IS NULL THEN '0' ELSE [04] END AS FERIADO, " + CRLF
cQuery += " CASE WHEN [05] IS NULL THEN '0' ELSE [05] END AS DSR " + CRLF
 
cQuery += " FROM (SELECT RC_FILIAL, RC_MAT,RA_NOME,RA_CIC,RC_CC,CTT_DESC01,RA_SITFOLH,RJ_DESC,RA_CATFUNC, RV_XRELHE AS TIPO,SUM(RC_VALOR) AS VALOR  FROM " +  retSqlname("SRC") + " SRC " + CRLF
		    
cQuery += " LEFT JOIN " + retSqlname("SRA") + " SRA " + CRLF
cQuery += " ON RA_FILIAL = RC_FILIAL AND RA_MAT = RC_MAT AND SRA.D_E_L_E_T_ = ''  " + CRLF
		    
cQuery += " LEFT JOIN " + retSqlname("SRV") + " SRV " + CRLF
cQuery += " ON RV_FILIAL = '" + xFilial("SRV") + "' AND RV_COD = RC_PD AND SRV.D_E_L_E_T_ = '' " + CRLF

cQuery += " LEFT JOIN " + retSqlname("CTT") + " CTT " + CRLF
cQuery += " ON CTT_FILIAL = '" + xFilial("CTT") + "' AND CTT_CUSTO = RC_CC AND CTT.D_E_L_E_T_ = '' " + CRLF

cQuery += " LEFT JOIN " + retSqlname("SRJ") + " SRJ " + CRLF
cQuery += " ON RJ_FILIAL = '" + xFilial("SRJ") + "' AND RA_CODFUNC = RJ_FUNCAO AND SRJ.D_E_L_E_T_ = '' " + CRLF
		    
cQuery += " WHERE SRC.D_E_L_E_T_ = '' " + CRLF
cQuery += " AND RC_PERIODO 		  = '" + MV_PAR02 + "' AND RC_ROTEIR = '" + MV_PAR01 + "' AND RC_SEMANA = '" + MV_PAR03 + "' " + CRLF
cQuery += " AND RC_FILIAL 	BETWEEN '" + MV_PAR04 + "' AND '" + MV_PAR05 + "' " + CRLF
cquery += " AND RC_MAT 		BETWEEN '" + MV_PAR06 + "' AND '" + MV_PAR07 + "' " + CRLF
cquery += " AND RC_CC 		BETWEEN '" + MV_PAR08 + "' AND '" + MV_PAR09 + "' " + CRLF
		    
cQuery += " GROUP BY RC_FILIAL,RC_MAT,RV_XRELHE,RA_NOME,RA_CIC,RC_CC,CTT_DESC01,RA_SITFOLH,RJ_DESC,RA_CATFUNC) P " + CRLF    
		    
cQuery += " PIVOT (SUM(VALOR) FOR TIPO IN ([00],[01],[02],[03],[04],[05])) AS PVT ), " + CRLF

cQuery += " CTE1 AS (SELECT " + CRLF
cQuery += " RD_FILIAL AS FILIALH,  " + CRLF
cQuery += " RD_MAT AS MATRICULAH,  " + CRLF

cQuery += " CASE WHEN [01] IS NULL THEN '0' ELSE [01] END AS HEH,  " + CRLF
cQuery += " CASE WHEN [02] IS NULL THEN '0' ELSE [02] END AS HENH,  " + CRLF
cQuery += " CASE WHEN [03] IS NULL THEN '0' ELSE [03] END AS BDH,  " + CRLF
cQuery += " CASE WHEN [04] IS NULL THEN '0' ELSE [04] END AS FERIADOH,  " + CRLF
cQuery += " CASE WHEN [05] IS NULL THEN '0' ELSE [05] END AS DSRH  " + CRLF

cQuery += " FROM (SELECT RD_FILIAL,RD_MAT,RV_XRELHE AS TIPO,SUM(RD_HORAS) AS HORAS  FROM " + retSqlname("SRD") + " SRD1  " + CRLF

cQuery += " LEFT JOIN " + retSqlname("SRV")+ " SRV1 " + CRLF
cQuery += " ON RV_FILIAL = '" + xFilial("SRV") + "' AND RV_COD = RD_PD AND SRV1.D_E_L_E_T_ = ''  " + CRLF

cQuery += " WHERE SRD1.D_E_L_E_T_ = '' "+ CRLF
cQuery += " AND RD_PERIODO 		  = '" + MV_PAR02 + "' AND RD_ROTEIR = '" + MV_PAR01 + "' AND RD_SEMANA = '" + MV_PAR03 + "' " + CRLF
cQuery += " AND RD_FILIAL 	BETWEEN '" + MV_PAR04 + "' AND '" + MV_PAR05 + "' " + CRLF
cquery += " AND RD_MAT 		BETWEEN '" + MV_PAR06 + "' AND '" + MV_PAR07 + "' " + CRLF
cquery += " AND RD_CC 		BETWEEN '" + MV_PAR08 + "' AND '" + MV_PAR09 + "' " + CRLF

cQuery += " GROUP BY RD_FILIAL,RD_MAT,RV_XRELHE) P   " + CRLF

cQuery += " PIVOT (SUM(HORAS) FOR TIPO IN ([00],[01],[02],[03],[04],[05])) AS PVT " + CRLF

cQuery += " UNION ALL " + CRLF

cQuery += " SELECT " + CRLF
cQuery += " RC_FILIAL AS FILIALH,  " + CRLF
cQuery += " RC_MAT AS MATRICULAH,  " + CRLF

cQuery += " CASE WHEN [01] IS NULL THEN '0' ELSE [01] END AS HEH,  " + CRLF
cQuery += " CASE WHEN [02] IS NULL THEN '0' ELSE [02] END AS HENH,  " + CRLF
cQuery += " CASE WHEN [03] IS NULL THEN '0' ELSE [03] END AS BDH,  " + CRLF
cQuery += " CASE WHEN [04] IS NULL THEN '0' ELSE [04] END AS FERIADOH,  " + CRLF
cQuery += " CASE WHEN [05] IS NULL THEN '0' ELSE [05] END AS DSRH  " + CRLF

cQuery += " FROM (SELECT RC_FILIAL,RC_MAT,RV_XRELHE AS TIPO,SUM(RC_HORAS) AS HORAS  FROM " + retSqlname("SRC") + " SRC1  " + CRLF

cQuery += " LEFT JOIN " + retSqlname("SRV")+ " SRV1 " + CRLF
cQuery += " ON RV_FILIAL = '" + xFilial("SRV") + "' AND RV_COD = RC_PD AND SRV1.D_E_L_E_T_ = ''  " + CRLF

cQuery += " WHERE SRC1.D_E_L_E_T_ = '' " + CRLF
cQuery += " AND RC_PERIODO 		  = '" + MV_PAR02 + "' AND RC_ROTEIR = '" + MV_PAR01 + "' AND RC_SEMANA = '" + MV_PAR03 + "' " + CRLF
cQuery += " AND RC_FILIAL 	BETWEEN '" + MV_PAR04 + "' AND '" + MV_PAR05 + "' " + CRLF
cquery += " AND RC_MAT 		BETWEEN '" + MV_PAR06 + "' AND '" + MV_PAR07 + "' " + CRLF
cquery += " AND RC_CC 		BETWEEN '" + MV_PAR08 + "' AND '" + MV_PAR09 + "' " + CRLF

cQuery += " GROUP BY RC_FILIAL,RC_MAT,RV_XRELHE) P " + CRLF

cQuery += " PIVOT (SUM(HORAS) FOR TIPO IN ([00],[01],[02],[03],[04],[05])) AS PVT) " + CRLF
		    
cQuery += " SELECT FILIAL, MATRICULA,NOME,CPF,CC,DESCCC,SITUACAO,FUNCAO,PESSOA, " + CRLF
cQuery += " SUM(HEH) AS HEH,SUM(HE) AS HE,SUM(HENH) AS HENH, SUM(HEN) AS HEN, SUM(BDH) AS BDH, " + CRLF
cQuery += " SUM(BH) AS BH,SUM(FERIADOH) AS FERIADOH,SUM(FERIADO) AS FERIADO,SUM(DSR) AS DSR, " + CRLF
cQuery += " SUM(HEH+HENH+BDH+FERIADOH) AS TOTALH,SUM(HE+HEN+BH+FERIADO+DSR) AS TOTALV FROM CTE " + CRLF  
cQuery += " LEFT JOIN CTE1 " + CRLF
cQuery += " ON FILIAL = FILIALH AND MATRICULA = MATRICULAH " + CRLF
cQuery += " GROUP BY FILIAL, MATRICULA,NOME,CPF,CC,DESCCC,SITUACAO,FUNCAO,PESSOA " + CRLF
cQuery += " HAVING SUM(HEH+HENH+BDH+FERIADOH) + SUM(HE+HEN+BH+FERIADO+DSR) > 0 "  + CRLF
cQuery += " ORDER BY FILIAL, NOME "

//cQuery := ChangeQuery(cQuery) //nao pode ser utilizada
DbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),cNextAlias)

Count To nCount
(cNextAlias)->(dbGoTop())
oReport:SetMeter(nCount)

oSection:Init()

While !(cNextAlias)->(Eof())

	oReport:IncMeter()
	oSection:PrintLine()
	//Cancelamento do relatório
	If oReport:Cancel()

		Exit

	EndIf	

	(cNextAlias)->(DbSkip())

EndDo

Return

/*------------------------------------------------------------------------------
|Programa: 		XRELHE															| 
|Tipo: 			Relatorio de Custos Horas Extras								|
|Empresa: 		Habib's 														|
|Analista: 		Bruno															|
|Consultoria:	Ethosx															|
-------------------------------------------------------------------------------*/

Static Function ValidPerg(cPerg)

Local aAlias := GetArea()
Local aRegs := {}
Local i,j

//
cPerg := PadR(cPerg, Len(SX1->X1_GRUPO), " ")
// 

aAdd(aRegs,{cPerg, "01", "Roteiro :		 	   ","","","mv_ch1","C",3,0,0,"G","","MV_PAR01","","","","","","","","","","","","","","","","","","","","","","","","","SRY","","",".RHROT.","",""})
aAdd(aRegs,{cPerg, "02", "Periodo :		 	   ","","","mv_ch2","C",6,0,0,"G","","MV_PAR02","","","","","","","","","","","","","","","","","","","","","","","","","RCH","","",".RHPER.","",""})
aAdd(aRegs,{cPerg, "03", "Numero de Pagamento: ","","","mv_ch3","C",2,0,0,"G","","MV_PAR03","","","","","","","","","","","","","","","","","","","","","","","","","","","",".RHNPA.","",""})
aAdd(aRegs,{cPerg, "04", "Filial de:	 	   ","","","mv_ch4","C",10,0,0,"G","","MV_PAR04","","","","","","","","","","","","","","","","","","","","","","","","","XM0","","",".RHFILDE.","",""})
aAdd(aRegs,{cPerg, "05", "Filial até:	 	   ","","","mv_ch5","C",10,0,0,"G","","MV_PAR05","","","","","","","","","","","","","","","","","","","","","","","","","XM0","","",".RHFILAT.","",""})
aAdd(aRegs,{cPerg, "06", "Matricula de:	 	   ","","","mv_ch6","C",6,0,0,"G","","MV_PAR06","","","","","","","","","","","","","","","","","","","","","","","","","SRA","","",".RHMATRIC.","",""})
aAdd(aRegs,{cPerg, "07", "Matricula até:	   ","","","mv_ch7","C",6,0,0,"G","","MV_PAR07","","","","","","","","","","","","","","","","","","","","","","","","","SRA","","",".RHMATRIC.","",""})
aAdd(aRegs,{cPerg, "08", "Centro de Custo de:  ","","","mv_ch8","C",10,0,0,"G","","MV_PAR08","","","","","","","","","","","","","","","","","","","","","","","","","CTT","","",".RHCCUSTO.","",""})
aAdd(aRegs,{cPerg, "09", "Centro de Custo até: ","","","mv_ch9","C",10,0,0,"G","","MV_PAR09","","","","","","","","","","","","","","","","","","","","","","","","","CTT","","",".RHCCUSTO.","",""})

//
DbSelectArea("SX1")                  
DbSetOrder(1)
For i := 1 to Len(aRegs)
	If !DbSeek(cPerg+aRegs[i,2])
		RecLock("SX1",.T.)
		For j := 1 to FCount()
			FieldPut(j,aRegs[i,j])
		Next
		MsUnlock()
	Endif
Next
RestArea( aAlias )

Return


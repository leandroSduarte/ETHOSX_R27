#INCLUDE "PROTHEUS.CH"
#INCLUDE "TBICONN.CH"
#INCLUDE "TOPCONN.CH"

/*
=====================================================================================
Programa.:
Autor....:              Felipe Ortega
Data.....:              31/01/2018
Descricao / Objetivo:   Bloquear os pedidos de compras cujo a entrega esteja em atraso
Doc. Origem:
Solicitante:            Cliente
Uso......:              HABIBS
Obs......:
=====================================================================================
*/

//-------------------------------------------------------------------
// Chamada via Schedule
user function FWSC23(aParams)

	local cEmp
	local cFilSch
	default aParams	:= {"01", "2502"}

	cEmp	:= aParams[1]
	cFilSch	:= aParams[2]

	If Type('cFilAnt')=='U'
		RpcSetType( 3 )
		RpcSetEnv( cEmp , cFilSch )
	Endif

	conout('[FWSC23] Iniciada Threads para a empresa 01  - ' + dToC(dDataBase) + " - " + time())

	QWSC23()



return
//-------------------------------------------------------------------
/*/{Protheus.doc} QWSC23
Rotina para listar a tabela de pedido de compras com os dados de
Vencimento. Caso esteja vencido o pedido e esteja dentro da filial 
sera liberado! caso contrario sera bloqueado 
@author  Leandro Duarte
@since   29/10/2021
@version 12.1.27
/*/
//-------------------------------------------------------------------
static function QWSC23()

	local aAreaSC7 	:= SC7->(GetArea())
	local aAreaZDX 	:= ZDX->(GetArea())
	local nDias    	:= SuperGetMV('FS_PRZENTR',,2)
	local cFilExc  	:= SuperGetMV('FS_FILEXC',,"2550")//Filiais que nÃ£o devem ter os pedidos excluidos/Bloquados.
	local nDiasExc	:= 0
	local cTab		:= getNextAlias()
	//by leandro Duarte
	IF XFILIAL("SC7") >= '2551' .and. (XFILIAL("SC7")<>'9000' .or. XFILIAL("SC7") <> '9001')
		nDiasExc    	:= SuperGetMV('FS_PRZENT2',,8)
	ENDIF
	cFilExc	:=	'%'+FormatIn(cFilExc,",")+'%'

	beginSql alias cTab
		SELECT
			C7_FILIAL,
			C7_NUM,
			C7_ITEM,
			C7_FORNECE,
			C7_LOJA,
			C7_DATPRF,
			SC7.R_E_C_N_O_ AS RECSC7
		FROM
			%Table:SC7% SC7
		WHERE
			C7_FILIAL NOT IN %exp:cFilExc%
			AND C7_QUJE <> C7_QUANT
			AND C7_CONAPRO = 'L'
			AND C7_ENCER <> 'E'
			AND C7_DATPRF <= %exp:DTOS(DDATABASE)%
			AND SC7.D_E_L_E_T_ = ''
	endSql
	/*Depois by Leandro Duarte*/
	dbselectarea(cTab)
	ZDX->(dbSetOrder(1))
	while (cTab)->(!eof())
		SC7->(DBGOTO((cTab)->RECSC7))
		If ZDX->(dbSeek(xFilial("ZDX") + (cTab)->C7_FORNECE + (cTab)->C7_LOJA))
			nDiasExc := ZDX->ZDX_DATLIM
			If Stod((cTab)->C7_DATPRF)+nDiasExc  < DDATABASE + nDias
				RecLock("SC7", .f.)
				SC7->C7_CONAPRO := 'B'
				SC7->C7_XBLPRF  := '1'
				SC7->C7_ENCER 	:= 'E'
				SC7->C7_RESIDUO := 'S'
				SC7->(msunlock())
			else // tratamento da rotina para caso tenha sido alterado o cadastro da ZDX by Leandro Duarte
				RecLock("SC7", .f.)
				SC7->C7_CONAPRO := 'L'
				SC7->C7_XBLPRF  := '2'
				SC7->C7_ENCER 	:= ' '
				SC7->C7_RESIDUO := ' '
				SC7->(msunlock())
			endif
		ELSE
			If Stod((cTab)->C7_DATPRF)+nDiasExc  < DDATABASE + nDias
				RecLock("SC7", .f.)
				SC7->C7_CONAPRO := 'B'
				SC7->C7_XBLPRF  := '1'
				SC7->C7_ENCER 	:= 'E'
				SC7->C7_RESIDUO := 'S'
				SC7->(msunlock())
			else // tratamento da rotina para caso tenha sido alterado o cadastro da ZDX by Leandro Duarte
				RecLock("SC7", .f.)
				SC7->C7_CONAPRO := 'L'
				SC7->C7_XBLPRF  := '2'
				SC7->C7_ENCER 	:= ' '
				SC7->C7_RESIDUO := ' '
				SC7->(msunlock())
			endif
		ENDIF
		(cTab)->(dbSkip())
	end
	ZDX->(dbCloseArea())
	SC7->(dbCloseArea())
	(cTab)->(dbCloseArea())
	restArea(aAreaZDX)
	restArea(aAreaSC7)
	// tratamento para casos quando a query não pegou pois estão bloqueados injustamente by leandro duarte
	beginSql alias cTab
		SELECT
			C7_FILIAL,
			C7_NUM,
			C7_ITEM,
			C7_FORNECE,
			C7_LOJA,
			C7_DATPRF,
			SC7.R_E_C_N_O_ AS RECSC7
		FROM
			%Table:SC7% SC7
		WHERE
			C7_FILIAL IN %exp:cFilExc%
			AND C7_QUJE <> C7_QUANT
			AND C7_CONAPRO = 'B'
			AND C7_ENCER = 'E'
			AND SC7.D_E_L_E_T_ = ''
	endSql
	/*Depois by Leandro Duarte*/
	dbselectarea(cTab)
	ZDX->(dbSetOrder(1))
	while (cTab)->(!eof())
		SC7->(DBGOTO((cTab)->RECSC7))
		If ZDX->(dbSeek(xFilial("ZDX") + (cTab)->C7_FORNECE + (cTab)->C7_LOJA))
			nDiasExc := ZDX->ZDX_DATLIM
			If Stod((cTab)->C7_DATPRF)+nDiasExc  >= DDATABASE + nDias
				RecLock("SC7", .f.)
				SC7->C7_CONAPRO := 'L'
				SC7->C7_XBLPRF  := '2'
				SC7->C7_ENCER 	:= ' '
				SC7->C7_RESIDUO := ' '
				SC7->(msunlock())
			endif
		ELSE
			If Stod((cTab)->C7_DATPRF)+nDiasExc >= DDATABASE + nDias
				RecLock("SC7", .f.)
				SC7->C7_CONAPRO := 'L'
				SC7->C7_XBLPRF  := '2'
				SC7->C7_ENCER 	:= ' '
				SC7->C7_RESIDUO := ' '
				SC7->(msunlock())
			endif
		ENDIF
		(cTab)->(dbSkip())
	end
	ZDX->(dbCloseArea())
	SC7->(dbCloseArea())
	(cTab)->(dbCloseArea())
	restArea(aAreaZDX)
	restArea(aAreaSC7)
Return

#INCLUDE "PROTHEUS.CH"

/*
Descrição

Ponto de entrada é executado quando gera título no financeiro e quando é efetuada a baixa do título.
*/

User Function MATATUSL()

	Local lRet   := SuperGetMV("MATATUSL",.F.,.F.)

Return(lRet)


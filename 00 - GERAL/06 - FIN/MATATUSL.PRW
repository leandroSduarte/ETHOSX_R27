#INCLUDE "PROTHEUS.CH"

/*
Descri��o

Ponto de entrada � executado quando gera t�tulo no financeiro e quando � efetuada a baixa do t�tulo.
*/

User Function MATATUSL()

	Local lRet   := SuperGetMV("MATATUSL",.F.,.F.)

Return(lRet)


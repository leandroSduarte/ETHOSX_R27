#INCLUDE "PROTHEUS.CH"

//-------------------------------------------------------------------
/*/{Protheus.doc} F040TRVSA1
O ponto de entrada F040TRVSA1 permite travar ou destravar os registros
da Tabela de Cliente - SA1, na rotina Clientes - MATA030. 
Essa ação é possível mesmo se os registros estiverem sendo utilizados por uma thread.

@author	Rafael tenorio da Costa 
@since 	16/09/16
@version 1.0
/*/
//-------------------------------------------------------------------
User Function F040TRVSA1() 

	Local lTrava := .T.
	
	// If IsInCallStack("LjGrvBatch")
		lTrava := .F.
	// EndIf
	 
Return lTrava
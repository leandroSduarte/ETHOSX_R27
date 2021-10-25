#Include 'Protheus.ch'

//------------------------------------------------------------------------------
/*/{Protheus.doc} HB06J02
description Rotina para controle das execu��es de baixas do
                 Contas a Pagar - Habib's'
@author  Ethosx - MOA
@since   01/09/2020
@version 1.0

@obs	Baseado na HB12J03 desenvolvida por Leonardo Espinosa
/*/
//-----------------------------------------------------------------------------
User Function HB06J02(cFilJob)

    Local cID
    Local nX

    Default cFilJob := "0001010001"

    cID := "06A02-"+cFilJob

    ManualJob(cID/*Nome do indentificador do job*/,;
			GetEnvServer()/*Ambiente que vc vai abrir este cara*/,;
			"IPC"/*Tipo do job. Mantenha como Ipc*/,;
			"U_06J02START"/*Fun��o que ser� chamada quando uma nova thread subir*/,;
			"U_06J02CONN" /*Fun��o que ser� chamada toda vez que vc mandar um ipcgo para ela*/,;
			"U_06J02EXIT"/*Fun��o que ser� invocada quando a thread cair pelo timeout dela*/,;
			cFilJob/*N�o alterar. � o SessionKey*/,;
			900/*Tempo que a thread ser� reavaliada e ir� cair. Vamos manter 5 minutos. Se n�o receber nada ela morre*/,;
			0/*Minimo de threads inicias. Vamos deixar 0 para que quando cair por timeout ele acabe*/,;
			1/*m�ximo de threads que ele vai subir*/,;
			1/*m�nimo de threads livres*/,;
			1/*incremento de threads livres*/)

        While !KillApp()
            
            // While !IpcGo(cID, .T., cFilJob, "000001","","1")
            //     Sleep(500)
            // EndDo

            IpcGo(cID, .T., cFilJob, "000001","","1")

            If !KillApp()
                For nX:=1 To 60
                    Sleep(1000)
                Next nX
            EndIf
        EndDo
        
Return

//-------------------------------------------------------------------
/*/{Protheus.doc} 06J02START
description Fun��o respons�vel por abrir o ambiente para integra��o
            das vendas.
@author  Leonardo Espinosa
@since   14/02/2020
@version 1.0
/*/
//-------------------------------------------------------------------
User Function 06J02START(cFilJob)

    RPCSetType(3)
    If RpcSetEnv('01',cFilJob,,,"FIN")
        ConOut("[06J02START - "+cFilJob+"] Ambiente aberto com sucesso!")
    EndIf 

Return .T.

//-------------------------------------------------------------------
/*/{Protheus.doc} 06J02CONN
description Fun��o que recebe o connect. Se chegou aqui, a de start
            j� chamou tamb�m.    
@author  Leonardo Espinosa
@since   14/02/2020
@version 1.0
/*/
//-------------------------------------------------------------------
User function 06J02CONN(lJob, cFiljob, cProc, cFilLinx,cFila)
    
    U_FINHAA02(lJob, cFilJob, cProc, cFilLinx, cFila)

Return 

//-------------------------------------------------------------------
/*/{Protheus.doc} 06J02EXIT
description Fun��o respons�vel por abrir o ambiente para integra��o
            das vendas.
@author  Leonardo Espinosa
@since   14/02/2020
@version 1.0
/*/
//-------------------------------------------------------------------
User Function 06J02EXIT(cFilJob)

	ConOut("[06J01EXIT - "+cFilJob+"] Ambiente finalizado pelo Timeout!")    
    RPCClearEnv( )

Return

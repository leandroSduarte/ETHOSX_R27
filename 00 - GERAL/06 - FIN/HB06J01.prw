#Include 'Protheus.ch'

//------------------------------------------------------------------------------
/*/{Protheus.doc} HB06J01
description Rotina para controle das execuções de baixas do
                 Contas a Receber - Habib's'
@author  Ethosx - MOA
@since   01/09/2020
@version 1.0

@obs	Baseado na HB12J03 desenvolvida por Leonardo Espinosa
/*/
//-----------------------------------------------------------------------------
User Function HB06J01(cFilJob)

    Local cID
    Local nX

    Default cFilJob := "0001010001"

    cID := "06A01-"+cFilJob

    ManualJob(cID/*Nome do indentificador do job*/,;
			GetEnvServer()/*Ambiente que vc vai abrir este cara*/,;
			"IPC"/*Tipo do job. Mantenha como Ipc*/,;
			"U_06J01START"/*Função que será chamada quando uma nova thread subir*/,;
			"U_06J01CONN" /*Função que será chamada toda vez que vc mandar um ipcgo para ela*/,;
			"U_06J01EXIT"/*Função que será invocada quando a thread cair pelo timeout dela*/,;
			cFilJob/*Não alterar. É o SessionKey*/,;
			900/*Tempo que a thread será reavaliada e irá cair. Vamos manter 5 minutos. Se não receber nada ela morre*/,;
			0/*Minimo de threads inicias. Vamos deixar 0 para que quando cair por timeout ele acabe*/,;
			1/*máximo de threads que ele vai subir*/,;
			1/*mínimo de threads livres*/,;
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
/*/{Protheus.doc} 06J01START
description Função responsável por abrir o ambiente para integração
            das vendas.
@author  Leonardo Espinosa
@since   14/02/2020
@version 1.0
/*/
//-------------------------------------------------------------------
User Function 06J01START(cFilJob)

    RPCSetType(3)
    If RpcSetEnv('01',cFilJob,,,"FIN")
        ConOut("[06J01START - "+cFilJob+"] Ambiente aberto com sucesso!")
    EndIf 

Return .T.

//-------------------------------------------------------------------
/*/{Protheus.doc} 06J01CONN
description Função que recebe o connect. Se chegou aqui, a de start
            já chamou também.    
@author  Leonardo Espinosa
@since   14/02/2020
@version 1.0
/*/
//-------------------------------------------------------------------
User function 06J01CONN(lJob, cFiljob, cProc, cFilLinx,cFila)
    
    U_HB06A01(lJob, cFilJob, cProc, cFilLinx, cFila)

Return 

//-------------------------------------------------------------------
/*/{Protheus.doc} 06J01EXIT
description Função responsável por abrir o ambiente para integração
            das vendas.
@author  Leonardo Espinosa
@since   14/02/2020
@version 1.0
/*/
//-------------------------------------------------------------------
User Function 06J01EXIT(cFilJob)

	ConOut("[06J01EXIT - "+cFilJob+"] Ambiente finalizado pelo Timeout!")    
    RPCClearEnv( )

Return

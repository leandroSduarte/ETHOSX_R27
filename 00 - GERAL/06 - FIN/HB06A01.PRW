#Include 'Protheus.ch'
#Include 'tbiconn.ch'

//-------------------------------------------------------------------
/*/{Protheus.doc} HB06A01
@description Rotina para efetuar baixas automaticas no
             contas a receber

@author  Leonardo Espinosa
@since   14/04/2020
@version 1.0
/*/
//-------------------------------------------------------------------
User Function HB06A01(aCompany)
Local   aRecs   := {  }
Local   aExec   := Array(08)

Local   cAlias  := ""
Local   cWhere  := ""
Local   cType   := ""
Local   cPrefix := ""
Local   cBank   := ""
Local   cAg     := ""
Local   cAccount:= ""
Local   cNature := ""
Local   cAux    := ""

Local   dBegin

Local   nX      := 0

Private cFunName    := "HB06A01"
Private lMSErroAuto := .F.

Default aCompany := {"01","0001010001"}

RPCSetType(3)
If RPCSetEnv(aCompany[01],aCompany[02],,,"FIN")
    ConOut("["+cFunName+"] Iniciando processamento... ["+FWTimeStamp(2)+"]")

    If SuperGetMV("MV_PRELAN") <> "S" .OR. SuperGetMV("MV_CT105MS") <> "N" .OR. SuperGetMV("MV_ALTLCTO") <> "N"
        ConOut("["+cFunName+"] Rotina parametrizada indevidamente. Abortando execucao....["+FWTimeStamp(2)+"]")
        ConOut("Somente será processada se: MV_PRELAN = S | MV_CT105MS = N | MV_ALTLCTO = N" )

        RPCClearEnv(  )

        Return
    EndIf

    cType   := AllTrim(SuperGetMV("HB_06A01TP",,"*"         ))
    cPrefix := AllTrim(SuperGetMV("HB_06A01PR",,"*"         ))
    cNature := AllTrim(SuperGetMV("HB_06A01NT",,"11110006"  ))
    cBank   := AllTrim(SuperGetMV("HB_06A01BC",,"CX01"      ))
    cAg     := AllTrim(SuperGetMV("HB_06A01AG",,"00001"     ))
    cAccount:= AllTrim(SuperGetMV("HB_06A01AC",,"0000000001"))

    dBegin  := STOD(SuperGetMV("HB_06A01BG",,"20200101"))

    cWhere := "%"
    If cType <> "*"
        cWhere += " AND E1_TIPO IN "+FormatIn(cType,";")
    Else
        cWhere += " AND E1_TIPO <> ('RA')"
    EndIf 

    If cPrefix <> "*"
        cWhere += " AND E1_PREFIXO IN "+FormatIn(cPrefix,";")
    EndIf

    cWhere += "AND E1_EMISSAO >= '"+DTOS(dBegin)+"'%"

    cAlias := GetNextAlias(  )
    If Select(cAlias) > 0
        (cAlias)->(DBCloseArea(  ))
    EndIf

    BeginSQL Alias cAlias
        SELECT
            SE1.R_E_C_N_O_ as RECSE1,
            E1_FILIAL,E1_PREFIXO,E1_NUM, E1_PARCELA,
            E1_CLIENTE,E1_LOJA,E1_NOMCLI,E1_EMISSAO,
            E1_VENCTO,E1_VENCREA,E1_FILORIG,
            E1_VALOR,E1_SALDO
        FROM
            %Table:SE1% (NOLOCK) SE1
        WHERE
            E1_FILIAL = %xFilial:SE1%
            %Exp:cWhere%
            AND E1_SALDO > 0
            AND E1_VENCREA = %Exp:DTOS(ddataBase)%
            AND SE1.%NotDel%
    EndSQL

    DBSelectArea(cAlias)
    If (cAlias)->(EOF(  ))
        ConOut("["+cFunName+"] Nenhum titulo encontrado para realizar a baixa em lote... ["+FWTimeStamp(2)+"]")
        ConOut(GetLastQuery()[2])
    Else
        Do while (cAlias)->(!EOF(  ))
            aAdd(aRecs, (cAlias)->RECSE1)
            (cAlias)->(DBSkip(  ))

        End Do 
    EndIf

    If Len(aRecs) > 0
        aExec[1] := aRecs
        aExec[2] := cBank
        aExec[3] := cAg
        aExec[4] := ""
        aExec[5] := "" //Lote
        aExec[7] := cNature
        aExec[8] := Date(  )

        Begin Transaction
            MSExecAuto({|x,y| Fina110(x,y)},3,aExec)

            If lMSErroAuto
                DisarmTransaction(  )
                ConOut("["+cFunName+"] Ocorreu um problema ao realizar a baixa dos titulos em lote... ["+FWTimeStamp(2)+"]")
                ConOut(Mostraerro())

            Else
                ConOut("["+cFunName+"] "+CVALTOCHAR( Len(aRecs) )+" titulo(s) baixados com sucesso! Lote: "+SE5->E5_LOTE+"... ["+FWTimeStamp(2)+"]")
            EndIf 

        End Transaction
    EndIf 

    ConOut("["+cFunName+"] Fim processamento... ["+FWTimeStamp(2)+"]")

    RPCClearEnv(  )
EndIf

Return
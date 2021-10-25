#INCLUDE "PROTHEUS.CH"

User Function F070TRAVA()

Local bRet := .T.

If IsInCallStack("U_HBFIN03F") .or. IsInCallStack("U_HBFIN03A") .or. IsInCallStack("U_HBFIN031") .or. IsInCallStack("U_HBFIN03H")
    bRet := .F.
EndIF

Return bRet

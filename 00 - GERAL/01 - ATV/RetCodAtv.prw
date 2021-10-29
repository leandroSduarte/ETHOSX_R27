#include "rwmake.ch"
#include "protheus.ch"
#include "colors.ch"

/*
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
�������������������������������������������������������������������������ͻ��
���Programa  �SeqProd   �Autor  �Carlos R Moreira    � Data �  09/20/09   ���
�������������������������������������������������������������������������͹��
���Desc.     �Gera o sequencial do codigo do Produto                      ���
���          �        teste                                               ���
�������������������������������������������������������������������������͹��
���Uso       � Especifico Dovac                                           ���
�������������������������������������������������������������������������ͼ��
�����������������������������������������������������������������������������
�����������������������������������������������������������������������������
*/

User Function RetCodAtv()
Local aArea := GetArea()
Local cCod := Space(10)
local cGrupo := M->N1_GRUPO
Local cCodIni := Space(4)

DbSelectArea("SNG")
DbSetOrder(1)
If DbSeek(xFilial("SNG")+M->N1_GRUPO)
   cCodIni := SNG->NG_CODINI
Else
   MsgStop("Grupo nao cadastrado")
   Return cCod 
EndIf     

BeginSql Alias "QRY"

  Select MAX(N1_CBASE) N1_CBASE FROM %Table:SN1% WHERE D_E_L_E_T_ <> '*'  AND N1_GRUPO = %EXP:cGrupo% AND
         SUBSTRING(N1_CBASE,1,4) = %Exp:cCodIni%   
      
EndSql 

DbSelectArea("QRY")
DbGoTop()

If Eof()
   
   cCodIni := cCodIni+"000001"
   QRY->(DbCloseArea())
   RestArea(aArea)
   Return cCod   
Else
   cCod := cCodIni+StrZero(Val(Substr(QRY->N1_CBASE,5,6)) + 1,6 )
EndIf 

QRY->(DbCloseArea())
   
RestArea(aArea)
Return cCod 

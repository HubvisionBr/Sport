/*/{Protheus.doc} nomeFunction
(long_description)
@type user function
@author user
@since 23/03/2026
@version version
@param param_name, param_type, param_descr
@return return_var, return_type, return_description
@example
(examples)
@see (links_or_references)
/*/

#include "Protheus.ch"

User Function MT235G1()

    Local cNumPed   := SC7->C7_NUM
    Local cCarteira := "P"
    Local cChaveFIE := xFilial("FIE") + cCarteira + cNumPed
    Local cChaveFR3 := xFilial("FR3") + cCarteira + cNumPed  


            FR3->(DbSetOrder(4))
            FR3->(DbSeek(cChaveFR3))
            While !FR3->(EoF())                      .And. ;
                FR3->FR3_FILIAL  == xFilial("FR3") .And. ;
                FR3->FR3_CART == cCarteira       .And. ;
                FR3->FR3_PEDIDO  == cNumPed

                RecLock("FR3", .F.)
                FR3->(DbDelete())
                MsUnlock()
                FR3->(DbSkip())
            EndDo

            FIE->(DbSetOrder(1))
            FIE->(DbSeek(cChaveFIE))
            While !FIE->(EoF())                      .And. ;
                  FIE->FIE_FILIAL  == xFilial("FIE") .And. ;
                  FIE->FIE_CART == cCarteira       .And. ;
                  FIE->FIE_PEDIDO  == cNumPed

                RecLock("FIE", .F.)
                FIE->(DbDelete())
                MsUnlock()
                FIE->(DbSkip())
            EndDo



Return

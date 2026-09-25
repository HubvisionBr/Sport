#Include "Protheus.ch"

/*/{Protheus.doc} XFinBlqBco
Funcao generica de validacao da data de liberacao do banco (SA6->A6_XDTBLOQ).
O banco/agencia/conta so aceita movimentacoes com data igual ou posterior
a A6_XDTBLOQ. Campo vazio = sem restricao.

Utilizada pelos pontos de entrada:
 - FINA070 : FA070TIT  (baixa a receber: manual, ExecAuto e lote)
 - FINA080 : FA080TIT  (baixa a pagar: manual, ExecAuto e lote)
 - FINA090 : F090TOK   (baixa automatica a pagar: tela e ExecAuto)
 - FINA100 : F100TOK   (movimento bancario a pagar/receber)
             FA100TRF  (transferencia bancaria e estorno de transferencia)

@param cBco     , character, Codigo do banco
@param cAge     , character, Agencia
@param cCta     , character, Conta
@param dData    , date     , Data da movimentacao a validar
@param cRotina  , character, Nome do PE/rotina (titulo do Help)
@param lHelp    , logical  , Exibe o Help em caso de bloqueio (default .T.)
@param cMsgErro , character, (referencia) retorna a mensagem de erro
@return lRet    , logical  , .T. permite / .F. bloqueia
/*/
User Function XFinBlqBco(cBco, cAge, cCta, dData, cRotina, lHelp, cMsgErro)
    Local lRet     := .T.
    Local aArea    := GetArea()
    Local aAreaSA6 := SA6->(GetArea())
    Local dDtLib   := CToD("")

    Default cBco     := ""
    Default cAge     := ""
    Default cCta     := ""
    Default dData    := dDataBase
    Default cRotina  := "XFINBLQBCO"
    Default lHelp    := .T.

    cMsgErro := ""

    If !Empty(cBco) .And. ValType(dData) == "D" .And. !Empty(dData)

        cBco := PadR(cBco, TamSX3("A6_COD")[1])
        cAge := PadR(cAge, TamSX3("A6_AGENCIA")[1])
        cCta := PadR(cCta, TamSX3("A6_NUMCON")[1])

        SA6->(DbSetOrder(1)) // A6_FILIAL+A6_COD+A6_AGENCIA+A6_NUMCON
        If SA6->(MsSeek(xFilial("SA6") + cBco + cAge + cCta))

            dDtLib := SA6->A6_XDTBLOQ

            If !Empty(dDtLib) .And. dData < dDtLib
                lRet     := .F.
                cMsgErro := "O banco " + AllTrim(cBco) + " / " + AllTrim(cAge) + " / " + AllTrim(cCta) + ;
                            " so permite movimentacoes a partir de " + DToC(dDtLib) + ;
                            ". Data informada: " + DToC(dData) + "."

                If lHelp
                    Help( , , cRotina, , cMsgErro, 1, 0, , , , , , ;
                        {"Informe uma data igual ou posterior a " + DToC(dDtLib) + ;
                         " ou utilize outro banco."} )
                EndIf
            EndIf
        EndIf
    EndIf

    RestArea(aAreaSA6)
    RestArea(aArea)
Return lRet

#Include "Protheus.ch"

/*/{Protheus.doc} F100TOK
Ponto de entrada de validacao do movimento bancario (FINA100).
Chamado em Fa100Bco(), que faz parte do TudoOk do AxInclui/AxIncluiAuto
de Movimento a Pagar (fa100pag) e a Receber (fa100rec) - tela e ExecAuto.
Valida a data de liberacao do banco (A6_XDTBLOQ) via U_XFinBlqBco.

Campos de memoria: M->E5_BANCO, M->E5_AGENCIA, M->E5_CONTA, M->E5_DATA
@return lRet, logical, .T. permite / .F. bloqueia
/*/
User Function F100TOK()
    Local lRet := .T.

    If Type("M->E5_BANCO") == "C" .And. Type("M->E5_DATA") == "D"
        lRet := U_XFinBlqBco(M->E5_BANCO, M->E5_AGENCIA, M->E5_CONTA, M->E5_DATA, "F100TOK")
    EndIf

Return lRet

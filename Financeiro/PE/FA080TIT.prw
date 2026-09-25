#Include "Protheus.ch"

/*/{Protheus.doc} FA080TIT
Ponto de entrada na confirmacao da Baixa a Pagar (FINA080).
Chamado em fA080Tit() (tela e MsExecAuto) e em Fa080But() (baixa por lote).
Retorno .F.: tela volta para correcao / ExecAuto aborta a baixa.
Valida a data de liberacao do banco (A6_XDTBLOQ) via U_XFinBlqBco.

Variaveis privadas da FINA080: cBanco, cAgencia, cConta (FINA080 / FA080Lot)
                               dBaixa (fA080Tit / FA080Lot)
@return lRet, logical, .T. permite a baixa / .F. bloqueia
/*/
User Function FA080TIT()
    Local lRet := .T.

    If Type("cBanco") == "C" .And. Type("dBaixa") == "D"
        lRet := U_XFinBlqBco(cBanco, cAgencia, cConta, dBaixa, "FA080TIT")
    EndIf

Return lRet

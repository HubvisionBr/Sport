#Include "Protheus.ch"

/*/{Protheus.doc} FA070TIT
Ponto de entrada na confirmacao da Baixa a Receber (FINA070).
Chamado em fA070Tit() (tela e MsExecAuto) e em Fa070But() (baixa por lote).
Valida a data de liberacao do banco (A6_XDTBLOQ) via U_XFinBlqBco.

Variaveis privadas da FINA070: cBanco, cAgencia, cConta, dBaixa
@return lRet, logical, .T. permite a baixa / .F. bloqueia
/*/
User Function FA070TIT()
    Local lRet := .T.

    If Type("cBanco") == "C" .And. Type("dBaixa") == "D"
        lRet := U_XFinBlqBco(cBanco, cAgencia, cConta, dBaixa, "FA070TIT")
    EndIf

Return lRet

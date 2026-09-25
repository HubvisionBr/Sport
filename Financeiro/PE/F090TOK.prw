#Include "Protheus.ch"

/*/{Protheus.doc} F090TOK
Ponto de entrada na validacao da tela de parametros da Baixa Automatica
a Pagar (FINA090), chamado em F090VldBx() no OK da tela e na rotina
automatica (aTitulos). Retorno .F. impede o processamento.
Valida a data de liberacao do banco (A6_XDTBLOQ) via U_XFinBlqBco.

PARAMIXB:
 [1] cBcoDe    [2] cBcoAte   [3] cBco090  [4] cAge090  [5] cCta090
 [6] cCheq090  [7] cBord090I [8] cBord090F [9] dVencIni [10] dVencFim
 [11] nTipoBx  [12] oRadio
Data da baixa: variavel privada dBaixa (FinA090)
@return lRet, logical, .T. permite / .F. bloqueia
/*/
User Function F090TOK()
    Local lRet := .T.

    If ValType(PARAMIXB) == "A" .And. Len(PARAMIXB) >= 5 .And. Type("dBaixa") == "D"
        lRet := U_XFinBlqBco(PARAMIXB[3], PARAMIXB[4], PARAMIXB[5], dBaixa, "F090TOK")
    EndIf

Return lRet

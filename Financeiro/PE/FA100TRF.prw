#Include "Protheus.ch"

/*/{Protheus.doc} FA100TRF
Ponto de entrada antes da gravacao da Transferencia Bancaria (FINA100).
Chamado em fa100tran() (transferencia) e Fa100Est() (estorno de
transferencia). Retorno .F. nao grava a transferencia.
Valida a data de liberacao (A6_XDTBLOQ) dos bancos de ORIGEM e DESTINO
via U_XFinBlqBco. A transferencia grava FK5_DATA = dDataBase nos dois lados.

PARAMIXB:
 [1] cBcoOrig [2] cAgenOrig [3] cCtaOrig
 [4] cBcoDest [5] cAgenDest [6] cCtaDest ... [15] lEstorno ...
@return lRet, logical, .T. grava / .F. nao grava
/*/
User Function FA100TRF()
    Local lRet := .T.

    If ValType(PARAMIXB) == "A" .And. Len(PARAMIXB) >= 6
        // Banco de origem
        lRet := U_XFinBlqBco(PARAMIXB[1], PARAMIXB[2], PARAMIXB[3], dDataBase, "FA100TRF")

        // Banco de destino
        If lRet
            lRet := U_XFinBlqBco(PARAMIXB[4], PARAMIXB[5], PARAMIXB[6], dDataBase, "FA100TRF")
        EndIf
    EndIf

Return lRet

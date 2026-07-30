#Include 'totvs.ch'

User Function MTALCDOC()
    Local aDoc := PARAMIXB[1]//array, 1-Numero do documento, 2-Tipo de documento, 3-Valor do documento, 4-Código do aprovador, 5-Código do usuário, 6-Grupo do aprovador, 7-Aprovador superior, 8-Moeda do documento, 9-Taxa da Moeda, 10-Data de emissão do documento
    Local cDocumento := aDoc[1]
    Local cTipo:= aDoc[2]
    // Local dData := PARAMIXB[2]//Data de referência do documento
    Local nTpOp := PARAMIXB[3]//1-Inclusão do documento, 2-Transferência da alçada para o superior, 3-Exclusão do documento, 4-Aprovação do documento, 5-Estorno da aprovação, 6-Bloqueio manual
    Local cGrpIT := PARAMIXB[4]//Item do grupo Obs: Parâmetro somente terá conteúdo quando aprovação for entidade contábil e por item
    Local cGrupo := aDoc[6]//grupo //NP3-09/07/2022
    Local cAprov := aDoc[4]//Aprovador//NP3-09/07/2022
    Local lAprov := .F.

    // PUBLIC aAPRVWF := {}

    If cTipo $ "MD|IM|PC|IP" 
        // se for aprovação, não vindo do fluig, não faz nada.
        IF (IsInCallStack("RESTCALLWS")) .and.  nTpOp == 4 
            RETURN
        EndIf

        //Se vier da cotação não envia para o fluig o pedido
        If IsInCallStack("MATA161") .or. IsInCallStack("MATA160") 
            RETURN
        EndIf

        //se for Aprovaçãototvsheus, manda uma uma exclusão //se tipo 7(rejeiçao) exclui
        if nTpOp == 4 .or. nTpOp == 7 
            lAprov := .T.
        EndIf

        U_APROVAWF(cDocumento,nTpOp,cTipo,cGrpIT,lAprov,cGrupo)
    EndIf

Return

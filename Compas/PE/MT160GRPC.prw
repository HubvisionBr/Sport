#Include 'Protheus.ch'

/*/{Protheus.doc} MT160GRPC
Ponto de entrada para gravar campos customizados na SC7 
durante a geração do pedido via Análise de Cotação.
/*/
User Function MT160GRPC()
    Local aAreaSC1 := SC1->(GetArea())
    Local aAreaSC7 := SC7->(GetArea())

    // A tabela SC7 já estará posicionada no pedido recém-gerado
    // Verifica se o pedido possui uma solicitação de compras vinculada
    If !Empty(SC7->C7_NUMSC)
        
        // Posiciona na SC1 (Filial + Numero da Solicitacao + Item da Solicitacao)
        SC1->(dbSetOrder(1)) 
        If SC1->(dbSeek(xFilial("SC1") + SC7->C7_NUMSC + SC7->C7_ITEMSC))
            
            // Grava o campo customizado na SC7 com o valor da SC1
            RecLock("SC7", .F.)
            SC7->C7_XIDFLUI := SC1->C1_XCODFLU
            SC7->(MsUnlock())
            
        EndIf
    EndIf

    RestArea(aAreaSC7)
    RestArea(aAreaSC1)
Return

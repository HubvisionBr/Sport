/*
+-----------------------------------------------------------------------+
¦Programa  ¦MT100GE2  ¦ Autor ¦ Raphael NEves         ¦ Data ¦06.08.2026¦
¦----------+------------------------------------------------------------¦
¦Descricao ¦ Ponto de entrada para gravar centro de custo               ¦
¦----------+------------------------------------------------------------¦
¦ Uso      ¦ ESPECIFICO PARA SPORT RECIFE                               ¦
¦-----------------------------------------------------------------------¦
¦           ATUALIZACOES SOFRIDAS DESDE A CONSTRUCAO INICIAL            ¦
¦-----------------------------------------------------------------------¦
¦PROGRAMADOR ¦ DATA   ¦ MOTIVO DA ALTERACAO                             ¦
¦------------+--------+-------------------------------------------------¦
¦            ¦        ¦                                                 ¦
+-----------------------------------------------------------------------+
*/

User Function MT100GE2()
	local cChaveSD1 := ""
	local aAreaSD1 := {}
	Local cCusto := ""

	IF SE2->(FieldPos("E2_CCUSTO")) > 0

		aAreaSD1 := SD1->(GetArea())
		POSICIONE("SD1",1,xFILIAL("SF1")+SF1->(F1_DOC+F1_SERIE+F1_FORNECE+F1_LOJA),"D1_DOC")
		cChaveSD1 := SD1->(D1_FILIAL+D1_DOC+D1_SERIE+D1_FORNECE+D1_LOJA)

		IF SD1->(!EOF()) .AND. cChaveSD1 = SD1->(D1_FILIAL+D1_DOC+D1_SERIE+D1_FORNECE+D1_LOJA)
			cCusto := SD1->D1_CC
		EndIf

		RestArea(aAreaSD1)

		SE2->E2_CCUSTO := cCusto
	Endif

Return

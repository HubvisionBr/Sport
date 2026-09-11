#INCLUDE "TOTVS.CH"

User Function FA750BRW()
	Local aBotao := {}
	Local aArea := GetArea()

	// AAdd(aBotao, { ;
	// 	"Baixa via arquivo", ; // Título do menu
	// 	"U_HVSPT01()", ; // Função associada
	// 	0, 4, 0, NIL;
	// })
	RestArea(aArea)

Return aBotao

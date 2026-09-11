#Include 'Protheus.ch'
#Include 'FWMVCDEF.ch'
 
User Function F050ROT()
     
    Local aArea   := GetArea()
    Local aRotina := Paramixb[1] // Array contendo os botoes padrões da rotina.
 
    // Tratamento no array aRotina para adicionar novos botoes e retorno do novo array.
    //Aadd(aRotina, {"Baixa via arquivo", "U_HVSPT01()", 0, 4,} )
     
    RestArea(aArea)
 
Return aRotina

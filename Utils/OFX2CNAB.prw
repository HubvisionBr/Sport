
#Include "Protheus.ch"

#Define PERG_OFX2CNAB  "OFX2CNAB "  


#Define DEF_BANCO      "125"
#Define DEF_AGENCIA    "00001"
#Define DEF_CONTA      "000001049230"
#Define DEF_CNPJ       "10866051000154"
#Define DEF_NOME_EMP   "SPORT CLUB DO RECIFE"
#Define DEF_NOME_BANCO "GENIAL INVESTIMENTOS"
#Define DEF_CONVENIO   ""


Static cST_BANCO      := DEF_BANCO
Static cST_AGENCIA    := DEF_AGENCIA
Static cST_DV_AG      := " "
Static cST_CONTA      := DEF_CONTA
Static cST_DV_CONTA   := " "
Static cST_DV_AGCTA   := " "
Static cST_TIPO_INSC  := "2"           // 1=CPF  2=CNPJ
Static cST_CNPJ       := DEF_CNPJ
Static cST_NOME_EMP   := DEF_NOME_EMP
Static cST_NOME_BANCO := DEF_NOME_BANCO
Static cST_CONVENIO   := DEF_CONVENIO


User Function OFX2CNAB()

    Local aRegs    := {}
    Local cPerg    := PERG_OFX2CNAB
    Local cBanco   := ""
    Local cAgencia := ""
    Local cConta   := ""
    Local cCNPJ    := ""
    Local cNomEmp  := ""
    Local cNomBco  := ""
    Local cConv    := ""
    Local cArqOFX  := ""
    Local cArqOut  := ""
    Local cNomeBase := ""
    Local cDirOFX   := ""
    Local nBarra    := 0
    Local nK := Len(cArqOFX)


    fnCriaSx1(aRegs, cPerg)

    If ! Pergunte(cPerg, .T.)
        Return Nil
    EndIf


    cBanco   := AllTrim(MV_PAR01)
    cAgencia := AllTrim(MV_PAR02)
    cConta   := AllTrim(MV_PAR03)
    cCNPJ    := AllTrim(MV_PAR04)
    cNomEmp  := AllTrim(MV_PAR05)
    cNomBco  := AllTrim(MV_PAR06)
    cConv    := AllTrim(MV_PAR07)
    cArqOFX  := AllTrim(MV_PAR08) //C:\TEMP\extrato.ofx
    cDirOFX  := AllTrim(MV_PAR09) //\servidor\cnab\

 
    If Empty(cArqOFX)
        MsgAlert("Informe o caminho do arquivo OFX de entrada (PAR08).", "Aten��o")
        Return Nil
    EndIf
    If !File(cArqOFX)
        MsgAlert("Arquivo OFX n�o encontrado:" + Chr(13) + cArqOFX, "Erro")
        Return Nil
    EndIf
    If Empty(cBanco) .Or. Empty(cAgencia) .Or. Empty(cConta) .Or. Empty(cCNPJ)
        MsgAlert("Banco, Ag�ncia, Conta e CNPJ s�o obrigat�rios.", "Aten��o")
        Return Nil
    EndIf


    // nBarra := 0
    // Do While nK > 0
        
    //     If SubStr(cArqOFX, nK, 1) == "\" .Or. SubStr(cArqOFX, nK, 1) == "/"
    //         nBarra := nK
    //         Exit
    //     EndIf
    //     nK--
    // EndDo

    nBarra := rat("\", cArqOFX)
    // cDirOFX   := If(nBarra > 0, Left(cArqOFX, nBarra), "")  
    cNomeBase := StrTran(SubStr(cArqOFX, nBarra + 1),'.','')

    ALERT(cDirOFX)


    If Empty(cArqOut)
      
        cArqOut := cDirOFX + cNomeBase + "_cnab240.txt"
        
    ElseIf Right(cArqOut, 1) == "\" .Or. Right(cArqOut, 1) == "/"
      
        cArqOut := cArqOut + cNomeBase + "_cnab240.txt"
    EndIf

    ALERT(cArqOut)

    cST_BANCO      := PadL(Left(cBanco,    3),  3, "0")
    cST_AGENCIA    := PadL(Left(cAgencia,  5),  5, "0")
    cST_CONTA      := PadL(Left(cConta,   12), 12, "0")
    cST_CNPJ       := PadL(Left(cCNPJ,   14), 14, "0")
    cST_NOME_EMP   := cNomEmp
    cST_NOME_BANCO := cNomBco
    cST_CONVENIO   := cConv


    OFXDoConv(cArqOFX, cArqOut)

Return Nil


Static Function fnCriaSx1(aRegs, cPerg)

    Local aAreaAtu := GetArea()
    Local aAreaSX1 := SX1->(GetArea())
    Local nJ       := 0
    Local nY       := 0


    aAdd(aRegs,{cPerg,"01","Banco (3 digitos)      ","","","mv_ch1","C",3,0,0,"G","","MV_PAR01","","","",DEF_BANCO,"","","","","","","","","","","","","","","","","","","","","","","","",""})
    aAdd(aRegs,{cPerg,"02","Agencia (5 digitos)    ","","","mv_ch2","C",5,0,0,"G","","MV_PAR02","","","",DEF_AGENCIA,"","","","","","","","","","","","","","","","","","","","","","","","",""})
    aAdd(aRegs,{cPerg,"03","Conta Corrente (12 dig)","","","mv_ch3","C",12,0,0,"G","","MV_PAR03","","","",DEF_CONTA,"","","","","","","","","","","","","","","","","","","","","","","","",""})
    aAdd(aRegs,{cPerg,"04","CNPJ (14 dig s/pontos) ","","","mv_ch4","C",14,0,0,"G","","MV_PAR04","","","",DEF_CNPJ,"","","","","","","","","","","","","","","","","","","","","","","","",""})
    aAdd(aRegs,{cPerg,"05","Nome da Empresa        ","","","mv_ch5","C",30,0,0,"G","","MV_PAR05","","","",DEF_NOME_EMP,"","","","","","","","","","","","","","","","","","","","","","","","",""})
    aAdd(aRegs,{cPerg,"06","Nome do Banco          ","","","mv_ch6","C",30,0,0,"G","","MV_PAR06","","","",DEF_NOME_BANCO,"","","","","","","","","","","","","","","","","","","","","","","","",""})
    aAdd(aRegs,{cPerg,"07","Convenio (deixe branco)","","","mv_ch7","C",20,0,0,"G","","MV_PAR07","","","",DEF_CONVENIO,"","","","","","","","","","","","","","","","","","","","","","","","",""})
    aAdd(aRegs,{cPerg,"08","Arquivo OFX (entrada)  ","","","mv_ch8","C",100,0,0,"G","","MV_PAR08","","","","","","","","","","","","","","","","","","","","","","","","","","","","",""})
    aAdd(aRegs,{cPerg,"09","Arq. CNAB saida(branco=auto)","","","mv_ch9","C",100,0,0,"G","","MV_PAR09","","","","","","","","","","","","","","","","","","","","","","","","","","","","",""})

    dbSelectArea("SX1")
    SX1->(dbSetOrder(1))

    For nY := 1 To Len(aRegs)
        If ! SX1->(dbSeek(Padr(cPerg, 10) + aRegs[nY, 2]))
            RecLock("SX1", .T.)
                For nJ := 1 To FCount()
                    If nJ <= Len(aRegs[nY])
                        FieldPut(nJ, aRegs[nY, nJ])
                    EndIf
                Next nJ
            SX1->(MsUnlock())
        EndIf
    Next nY

    RestArea(aAreaSX1)
    RestArea(aAreaAtu)

Return


Static Function OFXDoConv(cArqOFX, cArqOut)

    Local aTransacoes := {}
    Local cDateRef    := ""
    Local nBalFinal   := 0
    Local cDateCNAB   := ""
    Local cHoraCNAB   := ""
    Local nTotCre     := 0
    Local nTotDeb     := 0
    Local nSaldoIni   := 0
    Local nSaldoIniAbs:= 0
    Local cSitIni     := "C"
    Local cSitFin     := "C"
    Local nDet        := 0
    Local nLote       := 0
    Local nTotal      := 0
    Local aLinhas     := {}
    Local nHandle     := 0
    Local cLinha      := ""
    Local i
    Local cMsg        := ""
    Local cArqSrv     := ""


    If !ParseOFX(cArqOFX, @cDateRef, @nBalFinal, @aTransacoes)
        MsgAlert("Falha ao ler o arquivo OFX." + Chr(13) + ;
                 "Verifique se o arquivo está no formato correto.", "Erro")
        Return Nil
    EndIf
    If Len(aTransacoes) == 0
        MsgAlert("Nenhuma transação encontrada no arquivo OFX.", "Atenção")
        Return Nil
    EndIf

    cDateCNAB := DDMMAAAA(cDateRef)
    cHoraCNAB := StrTran(Time(), ":", "")    


    For i := 1 To Len(aTransacoes)
        If aTransacoes[i][3] == "C"
            nTotCre += aTransacoes[i][2]
        Else
            nTotDeb += aTransacoes[i][2]
        EndIf
    Next i


    nSaldoIni    := nBalFinal - nTotCre + nTotDeb
    cSitIni      := If(nSaldoIni   >= 0, "C", "D")
    cSitFin      := If(nBalFinal   >= 0, "C", "D")
    nSaldoIniAbs := Abs(nSaldoIni)
    nBalFinal    := Abs(nBalFinal)

  
    nDet   := Len(aTransacoes)
    nLote  := nDet + 2     
    nTotal := nLote + 2   

    //-- Monta registros
    AAdd(aLinhas, HdrArquivo(cDateCNAB, cHoraCNAB, "000001"))
    AAdd(aLinhas, HdrLote(cDateCNAB, nSaldoIniAbs, cSitIni))
    For i := 1 To nDet
        AAdd(aLinhas, DetalheE(i, aTransacoes[i]))
    Next i
    AAdd(aLinhas, TrlLote(nLote, cDateCNAB, nBalFinal, cSitFin, nTotDeb, nTotCre))
    AAdd(aLinhas, TrlArquivo(1, nTotal))



    //-- Grava direto no servidor com o caminho informado
    cArqSrv := cArqOut

    nHandle := FCreate(cArqSrv)
    If nHandle < 0
        MsgAlert("Nao foi possivel criar o arquivo no servidor:" + Chr(13) + cArqSrv + Chr(13) + Chr(13) + ;
                 "Verifique se o diretorio existe e se ha permissao de escrita.", "Erro de Escrita")
        Return Nil
    EndIf

    For i := 1 To Len(aLinhas)
        cLinha := aLinhas[i] + Chr(13) + Chr(10)
        FWrite(nHandle, cLinha, Len(cLinha))
    Next i
    FClose(nHandle)

    //-- Relatorio de resultado
    cMsg := "CNAB 240 gerado com sucesso!" + Chr(13) + Chr(13) + ;
            "Arquivo   : " + cArqSrv           + Chr(13) + ;
            "Extrato   : " + cDateCNAB          + Chr(13) + ;
            "Transacoes: " + AllTrim(Str(nDet)) + Chr(13) + ;
            "Sal. Inic.: " + Transform(nSaldoIniAbs, "@E 999,999,999,999.99") + " [" + cSitIni + "]" + Chr(13) + ;
            "Sal. Final: " + Transform(nBalFinal,    "@E 999,999,999,999.99") + " [" + cSitFin + "]" + Chr(13) + ;
            "Tot.Cred. : " + Transform(nTotCre,      "@E 999,999,999,999.99") + Chr(13) + ;
            "Tot.Deb.  : " + Transform(nTotDeb,      "@E 999,999,999,999.99") + Chr(13) + Chr(13) + ;
            "Reg. Lote : " + AllTrim(Str(nLote))  + "  (1 + " + AllTrim(Str(nDet)) + " det + 1)" + Chr(13) + ;
            "Reg. Total: " + AllTrim(Str(nTotal))

    MsgInfo(cMsg, "Conversao Concluida")

Return Nil


Static Function ParseOFX(cFile, cDateRef, nBalance, aTransacoes)

    Local nArq     := FOpen(cFile, 0)  
    Local nSize    := 0
    Local cContent := ""
    Local aLines   := {}
    Local cLine    := ""
    Local cUpper   := ""
    Local cVal     := ""
    Local lInTrn   := .F.
    Local cData    := ""
    Local nValor   := 0
    Local cTipo    := ""
    Local cCheck   := ""
    Local cMemo    := ""
    Local i, nPos

    aTransacoes := {}
    cDateRef    := DToS(Date())  
    nBalance    := 0

    If nArq < 0
        Return .F.
    EndIf

    nSize    := FSeek(nArq, 0, 2)   
    FSeek(nArq, 0, 0)               
    cContent := Space(nSize)
    FRead(nArq, @cContent, nSize)
    FClose(nArq)

  
    cContent := StrTran(cContent, Chr(13), "")   
    aLines   := SplitStr(cContent, Chr(10))      

    For i := 1 To Len(aLines)
     
        cLine  := AllTrim(StrTran(aLines[i], Chr(9), ""))
        cUpper := Upper(cLine)


        nPos := At("<BALAMT>", cUpper)
        If nPos > 0
            cVal     := AllTrim(SubStr(cLine, nPos + 8))
           
            nPos := At("[", cVal)
            If nPos > 0
                cVal := Left(cVal, nPos - 1)
            EndIf
            nBalance := Val(AllTrim(cVal))
        EndIf

        nPos := At("<DTSTART>", cUpper)
        If nPos > 0
            cVal     := AllTrim(SubStr(cLine, nPos + 9))
            cDateRef := Left(cVal, 8)
        EndIf

        If At("<STMTTRN>", cUpper) > 0 .And. At("</STMTTRN>", cUpper) == 0
            lInTrn := .T.
            cData  := ""
            nValor := 0
            cTipo  := "C"
            cCheck := ""
            cMemo  := ""
        EndIf


        If lInTrn

            nPos := At("<DTPOSTED>", cUpper)
            If nPos > 0
             
                cVal  := AllTrim(SubStr(cLine, nPos + 10))
                cData := Left(cVal, 8)
            EndIf

            nPos := At("<TRNAMT>", cUpper)
            If nPos > 0
                cVal   := AllTrim(SubStr(cLine, nPos + 8))
                nValor := Val(cVal)
                cTipo  := If(nValor >= 0, "C", "D")
                nValor := Abs(nValor)
            EndIf

            nPos := At("<CHECKNUM>", cUpper)
            If nPos > 0
                cCheck := AllTrim(SubStr(cLine, nPos + 10))
            EndIf

            nPos := At("<MEMO>", cUpper)
            If nPos > 0
                cMemo := AllTrim(SubStr(cLine, nPos + 6))
            EndIf

        EndIf

     
        If At("</STMTTRN>", cUpper) > 0
            If lInTrn .And. !Empty(cData)
                AAdd(aTransacoes, { cData, nValor, cTipo, ;
                                    Left(cCheck + Space(20), 20), ;
                                    Left(cMemo  + Space(25), 25) })
            EndIf
            lInTrn := .F.
        EndIf

    Next i

    Return .T.


Static Function CnabA(cVal, nSize)
    Local cR := Left(cVal + Space(nSize), nSize)
Return cR


Static Function CnabN(nVal, nSize)
    Local cR := StrZero(Int(nVal), nSize)
    If Len(cR) > nSize
        cR := Right(cR, nSize)  
    EndIf
Return cR


Static Function CnabNS(cVal, nSize)
    Local cR := PadL(Left(AllTrim(cVal), nSize), nSize, "0")
Return cR

Static Function CnabV(nAmount, nDig, nDec)
    Local nFator := 1
    Local nCents := 0
    Local nTotal := nDig + nDec
    Local cR

    Default nDig := 16
    Default nDec := 2


    nFator := iif(nDec == 2, 100, iif(nDec == 3, 1000, 10))

    nCents := Int(Abs(nAmount) * nFator + 0.5)   
    cR     := StrZero(nCents, nTotal)
    If Len(cR) > nTotal
        cR := Right(cR, nTotal)
    EndIf
Return cR


Static Function DDMMAAAA(cYMD)
    Local cS := Left(AllTrim(cYMD) + "00000000", 8)
Return SubStr(cS, 7, 2) + SubStr(cS, 5, 2) + Left(cS, 4)


Static Function SplitStr(cStr, cDelim)
    Local aRet := {}
    Local cRest := cStr
    Local nPos

    Do While .T.
        nPos := At(cDelim, cRest)
        If nPos == 0
            Exit
        EndIf
        AAdd(aRet, Left(cRest, nPos - 1))
        cRest := SubStr(cRest, nPos + Len(cDelim))
    EndDo
    If !Empty(cRest)
        AAdd(aRet, cRest)
    EndIf
Return aRet


Static Function FNameNoExt(cPath)
    Local nDot := OFXRAt(".", cPath)
    If nDot > 0
        Return Left(cPath, nDot - 1)
    EndIf
Return cPath


Static Function OFXRAt(cSub, cStr)
    Local nLast := 0
    Local nPos  := 0

    nPos := At(cSub, cStr)
    Do While nPos > 0
        nLast += nPos
        nPos   := At(cSub, SubStr(cStr, nLast + 1))
    EndDo
Return nLast




Static Function HdrArquivo(cDateCNAB, cHora, cNSA)
    Local r := ""

    Default cNSA := "000001"

    r += CnabN(Val(cST_BANCO), 3)          // 001-003  Banco na compensação
    r += "0000"                             // 004-007  Lote = "0000"
    r += "0"                                // 008-008  Tipo registro = "0"
    r += Space(9)                           // 009-017  CNAB (brancos)
    r += cST_TIPO_INSC                      // 018-018  Tipo inscrição 1=CPF / 2=CNPJ
    r += CnabNS(cST_CNPJ, 14)              // 019-032  Nº inscrição da empresa (14)
    r += CnabA(cST_CONVENIO, 20)            // 033-052  Código convênio banco (20 alfa)
    r += cST_AGENCIA                        // 053-057  Agência (5 num)
    r += cST_DV_AG                          // 058-058  DV agência (1 alfa)
    r += cST_CONTA                          // 059-070  Conta corrente (12 num)
    r += cST_DV_CONTA                       // 071-071  DV conta (1 alfa)
    r += cST_DV_AGCTA                       // 072-072  DV ag/conta (1 alfa)
    r += CnabA(cST_NOME_EMP,  30)           // 073-102  Nome empresa (30 alfa)
    r += CnabA(cST_NOME_BANCO, 30)          // 103-132  Nome banco (30 alfa)
    r += Space(10)                          // 133-142  CNAB (brancos)
    r += "2"                                // 143-143  Código retorno = "2"
    r += cDateCNAB                          // 144-151  Data geração DDMMAAAA (8 num)
    r += cHora                              // 152-157  Hora geração HHMMSS (6 num)
    r += PadL(Left(cNSA, 6), 6, "0")        // 158-163  Nº sequencial arquivo (6 num)
    r += "030"                              // 164-166  Versão layout = "030"
    r += "00000"                            // 167-171  Densidade de gravação (5 num)
    r += Space(20)                          // 172-191  Reservado banco (20 alfa)
    r += Space(20)                          // 192-211  Reservado empresa (20 alfa)
    r += Space(11)                          // 212-222  CNAB (brancos)
    r += "CSP"                              // 223-225  Cobrança s/papel
    r += "000"                              // 226-228  Controle VANs (3 num)
    r += "  "                               // 229-230  Tipo de serviço (2 alfa)
    r += Space(10)                          // 231-240  Código ocorrências (10 alfa)


    r := Left(r + Space(240), 240)
Return r


Static Function HdrLote(cDateCNAB, nSaldoIni, cSitIni)
    Local r := ""

    r += CnabN(Val(cST_BANCO), 3)          // 001-003
    r += "0001"                             // 004-007  Lote = "0001"
    r += "1"                                // 008-008  Tipo registro = "1"
    r += "E"                                // 009-009  Operação = "E" (extrato C/C)
    r += "06"                               // 010-011  Tipo serviço extrato
    r += "01"                               // 012-013  Forma lançamento
    r += "030"                              // 014-016  Versão layout lote
    r += " "                                // 017-017  CNAB
    r += cST_TIPO_INSC                      // 018-018
    r += CnabNS(cST_CNPJ, 14)              // 019-032
    r += CnabA(cST_CONVENIO, 20)            // 033-052
    r += cST_AGENCIA                        // 053-057  (5 num)
    r += cST_DV_AG                          // 058-058
    r += cST_CONTA                          // 059-070  (12 num)
    r += cST_DV_CONTA                       // 071-071
    r += cST_DV_AGCTA                       // 072-072
    r += CnabA(cST_NOME_EMP, 30)            // 073-102
    r += CnabA("", 40)                      // 103-142  Segunda linha extrato (40 alfa)
    r += cDateCNAB                          // 143-150  Data saldo inicial DDMMAAAA (8)
    r += CnabV(nSaldoIni, 16, 2)            // 151-168  Valor saldo inicial (18 chars)
    r += cSitIni                            // 169-169  D=devedor / C=credor
    r += "F"                                // 170-170  Posição = "F" (final)
    r += "BRL"                              // 171-173  Moeda
    r += CnabN(1, 5)                        // 174-178  Nº sequência extrato (5 num)
    r += Space(62)                          // 179-240  CNAB (brancos)

    r := Left(r + Space(240), 240)
Return r


Static Function DetalheE(nSeq, aTrn)
    Local r        := ""
    Local cCheck   := Left(aTrn[4] + Space(20), 20)
    Local cMemoTxt := Left(aTrn[5] + Space(25), 25)
    Local cDigits  := ""
    Local cCodBco  := "0000"
    Local i, cCh


    For i := 1 To Len(cCheck)
        cCh := SubStr(cCheck, i, 1)
        If cCh >= "0" .And. cCh <= "9"
            cDigits += cCh
        EndIf
    Next i

    If Len(cDigits) >= 4
        cCodBco := Right(cDigits, 4)
    ElseIf !Empty(cDigits)
        cCodBco := PadL(cDigits, 4, "0")
    EndIf

    r += CnabN(Val(cST_BANCO), 3)          // 001-003
    r += "0001"                             // 004-007
    r += "3"                                // 008-008  Tipo registro = "3"
    r += CnabN(nSeq, 5)                    // 009-013  Nº sequencial detalhe no lote
    r += "E"                                // 014-014  Segmento = "E"
    r += Space(3)                           // 015-017  CNAB (brancos)
    r += cST_TIPO_INSC                      // 018-018
    r += CnabNS(cST_CNPJ, 14)              // 019-032
    r += CnabA(cST_CONVENIO, 20)            // 033-052
    r += cST_AGENCIA                        // 053-057  (5 num)
    r += cST_DV_AG                          // 058-058
    r += cST_CONTA                          // 059-070  (12 num)
    r += cST_DV_CONTA                       // 071-071
    r += cST_DV_AGCTA                       // 072-072
    r += CnabA(cST_NOME_EMP, 30)            // 073-102
    r += Space(40)                          // 103-142  CNAB (brancos)
    r += DDMMAAAA(aTrn[1])                  // 143-150  Data lançamento DDMMAAAA (8)
    r += CnabV(aTrn[2], 16, 2)             // 151-168  Valor lançamento (18 chars)
    r += aTrn[3]                            // 169-169  D=débito / C=crédito
    r += If(aTrn[3] == "D", "002", "001")   // 170-172  Categoria lançamento (3 num)
    r += cCodBco                            // 173-176  Código lançamento banco (4 num)
    r += CnabA(cMemoTxt, 25)               // 177-201  Histórico (25 alfa)
    r += CnabA(cCheck,   20)               // 202-221  Nº documento comprobatório (20 alfa)
    r += Space(19)                          // 222-240  CNAB (brancos)

    r := Left(r + Space(240), 240)
Return r


Static Function TrlLote(nQtdRegs, cDateCNAB, nSaldoFin, cSitFin, nTotDeb, nTotCre)
    Local r := ""

    r += CnabN(Val(cST_BANCO), 3)          // 001-003
    r += "0001"                             // 004-007
    r += "5"                                // 008-008  Tipo registro = "5"
    r += Space(9)                           // 009-017  CNAB (brancos)
    r += cST_TIPO_INSC                      // 018-018
    r += CnabNS(cST_CNPJ, 14)              // 019-032
    r += CnabA(cST_CONVENIO, 20)            // 033-052
    r += cST_AGENCIA                        // 053-057  (5 num)
    r += cST_DV_AG                          // 058-058
    r += cST_CONTA                          // 059-070  (12 num)
    r += cST_DV_CONTA                       // 071-071
    r += cST_DV_AGCTA                       // 072-072
    r += CnabA(cST_NOME_EMP, 30)            // 073-102
    r += Space(4)                           // 103-106  CNAB (brancos)
    r += CnabV(0, 16, 2)                    // 107-124  Limite da conta (18 chars)
    r += CnabV(0, 16, 2)                    // 125-142  Saldo bloqueado (18 chars)
    r += cDateCNAB                          // 143-150  Data saldo final DDMMAAAA (8)
    r += CnabV(nSaldoFin, 16, 2)           // 151-168  Valor saldo final (18 chars)
    r += cSitFin                            // 169-169  D=devedor / C=credor
    r += "F"                                // 170-170  Posição = "F" (final)
    r += CnabN(nQtdRegs, 6)               // 171-176  Qtd registros do lote (6 num)
    r += CnabV(nTotDeb, 16, 2)            // 177-194  Somatória débitos (18 chars)
    r += CnabV(nTotCre, 16, 2)            // 195-212  Somatória créditos (18 chars)
    r += Space(28)                          // 213-240  CNAB (brancos)

    r := Left(r + Space(240), 240)
Return r


Static Function TrlArquivo(nQtdLotes, nQtdTotalRegs)
    Local r := ""

    r += CnabN(Val(cST_BANCO), 3)          // 001-003
    r += "9999"                             // 004-007  Lote = "9999"
    r += "9"                                // 008-008  Tipo registro = "9"
    r += Space(9)                           // 009-017  CNAB (brancos)
    r += CnabN(nQtdLotes,     6)           // 018-023  Qtd lotes do arquivo (6 num)
    r += CnabN(nQtdTotalRegs, 6)           // 024-029  Qtd registros do arquivo (6 num)
    r += CnabN(1, 6)                        // 030-035  Qtd contas conciliadas (6 num)
    r += Space(205)                         // 036-240  CNAB (brancos)

    r := Left(r + Space(240), 240)
Return r

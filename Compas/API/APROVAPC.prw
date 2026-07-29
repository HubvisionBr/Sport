#Include 'FWMVCDEF.ch'
#Include 'RestFul.CH'
#INCLUDE "TOTVS.CH"
#INCLUDE "TopConn.ch"
#INCLUDE "TBICONN.CH"
#Include 'parmtype.ch'

#DEFINE ENTER CHR(13)+CHR(10)

//Início da declaração da estrutura do Webservice;
WSRESTFUL APROVAPC DESCRIPTION "Aprovacao"

	WSMETHOD POST DESCRIPTION "Aprovacao de pedido de compra" WSSYNTAX " enviar via body dados conforme documentacao"
END WSRESTFUL

WSRESTFUL RETAPROV DESCRIPTION "Retorna email de aprovador "

	WSMETHOD POST DESCRIPTION "Retorna email de aprovador" WSSYNTAX " enviar via body dados conforme documentacao"
END WSRESTFUL

WSRESTFUL ENVAPROV DESCRIPTION "Envia aprovação "

	WSMETHOD POST DESCRIPTION "Envia aprovação" WSSYNTAX " enviar via body dados conforme documentacao"
END WSRESTFUL

WSRESTFUL APROVAMED DESCRIPTION "Aprova Medição "

	WSMETHOD POST DESCRIPTION "Aprova Medição" WSSYNTAX " enviar via body dados conforme documentacao"
END WSRESTFUL

WSRESTFUL RETSALDO DESCRIPTION "Retorna Saldo Conta Orçamentaria"

	WSMETHOD POST DESCRIPTION "Retorna Saldo Conta Orçamentaria" WSSYNTAX " enviar via body dados conforme documentacao"
END WSRESTFUL

WSMETHOD POST WSSERVICE RETAPROV
	Local cJson           := Self:GetContent()
	Local oJson
	Private cMensagem := ""

	::SetContentType("application/json")

	oJson := JsonObject():New()

	oRet := oJson:FromJson(cJson)

	if ValType(oRet) == "U"
		aDados := oJson:GetNames()
		nPosApro := aScan(aDados,{|x| x == "aprovador"})

		cCnpj := ""
		cAprov:= oJson:GetJsonText(aDados[nPosApro])

		//Verifica se o aprovador está ausente
		dbSelectArea("SAK")
		SAK->(dbSetOrder(1))
		IF SAK->(MsSeek(xFilial("SAK")+cAprov))
			__cUserID := SAK->AK_USER

			::SetResponse("{'nome':'"+Alltrim(FwGetUserName(RetCodUsr()))+"','email':'"+AllTrim(UsrRetMail(RetCodUsr()))+"'}")
		Else
			::SetResponse("{'nome':'nao_encontrado','email':''}")
		Endif
	Endif
	FreeObj(oJson)

Return .T.


WSMETHOD POST WSSERVICE APROVAPC

	Local cJson           := Self:GetContent()
	Local oJson
	Private cMensagem := ""

	// u_logInt("2", "APROVAPC", Self:GetContent())

	::SetContentType("application/json")

	oJson := JsonObject():New()

	oRet := oJson:FromJson(cJson)

	if ValType(oRet) == "U"
		IF Aprova(oJson)
			::SetResponse("{'status':'success','message':'pedido aprovado'}")
		Else
			::setStatus(400)
			::SetResponse("{'status':'error','message':'"+cMensagem+"'}")
		Endif
	else
		Conout("Falha ao popular JsonObject. Erro: " + oRet)
	endif
	FreeObj(oJson)

Return .T.


WSMETHOD POST WSSERVICE ENVAPROV

	Local cJson           := Self:GetContent()
	Local oJson
	Private cMensagem := ""

	// u_logInt("2", "ENVAPROV", Self:GetContent())

	::SetContentType("application/json")

	oJson := JsonObject():New()

	oRet := oJson:FromJson(cJson)

	if ValType(oRet) == "U"
		IF Aprova(oJson,.T.)
			::SetResponse("{'status':'success','message':'pedido aprovado'}")
		Else
			::setStatus(400)
			::SetResponse("{'status':'error','message':'"+cMensagem+"'}")
		Endif
	else
		Conout("Falha ao popular JsonObject. Erro: " + oRet)
	endif
	FreeObj(oJson)

Return .T.

WSMETHOD POST WSSERVICE APROVAMED

	Local cJson           := Self:GetContent()
	Local oJson
	Local oModel
	Local aDados
	Local cFilFlu,cMed,cContra,cStatus
	Local nFilFlu,nMed,nContra,nStatus
	Private cMensagem := ""

	// u_logInt("2", "APROVAMED", Self:GetContent())

	::SetContentType("application/json")

	oJson := JsonObject():New()

	oRet := oJson:FromJson(cJson)

	if ValType(oRet) == "U"
		aDados := oJson:GetNames()

		nFilFlu := aScan(aDados,{|x| x == "filial"})
		nMed 	:= aScan(aDados,{|x| x == "medicao"})
		nContra := aScan(aDados,{|x| x == "contrato"})
		nStatus := aScan(aDados,{|x| x == "status"})

		cFilFlu := AllTrim(oJson:GetJsonText(aDados[nFilFlu]))
		cMed 	:= oJson:GetJsonText(aDados[nMed])
		cContra := oJson:GetJsonText(aDados[nContra])
		cStatus := oJson:GetJsonText(aDados[nStatus])

		DbSelectArea("CND")
		CND->(DbSetOrder(4))//CND_FILIAL+CND_NUMMED
		If CND->(DbSeek(PadR(cFilFlu,TamSx3("CND_FILIAL")[1])+ PadR(cMed,TamSx3("CND_NUMMED")[1])))

			If AllTrim(cStatus) == "Aprovado"
				CN121Encerr(.T.)
				::SetResponse("{'status':'success','message':'medicao encerrada'}")
			Else

				oModel := FWLoadModel("CNTA121")

				oModel:SetOperation(MODEL_OPERATION_DELETE)
				If(oModel:CanActivate())
					oModel:Activate()
					If (oModel:VldData()) //Valida o modelo como um todo
						oModel:CommitData()
					EndIf
				EndIf

				::SetResponse("{'status':'success','message':'medicao excluida'}")
			EndIf
		Else
			::setStatus(400)
			::SetResponse("{'status':'error','message':'medicao nao encontrada'}")
		EndIf
	else
		Conout("Falha ao popular JsonObject. Erro: " + oRet)
		if ValType(cRet) == "C"
			SetRestFault(400,'Erro de Estrutura no Json')
		endif
	endif

	FreeObj(oJson)

Return .T.

/*******************************/
WSMETHOD POST WSSERVICE RETSALDO

	Local cJson           := Self:GetContent()
	Local oJson
	Local aDados    := {}
	Local cFilFlu   := ""
	Local cConta    := ""
	Local cData     := ""
	Local nFilFlu   := 0
	Local nConta    := 0
	Local nData     := 0
	Local nCCusto   := 0
	Local cCCusto   := 0
	Local nSaldo    := 0
	Local cEmail    := ""
	Local lRet      := .T.
	Local cError    := ""
	Local cErr      := ""

	Local oLastError := ErrorBlock( { |e| cErr := e:ErrorStack, Break(e) } )


	Begin Sequence
		::SetContentType("application/json")

		oJson := JsonObject():New()

		oRet := oJson:FromJson(cJson)

		if ValType(oRet) == "U"
			aDados := oJson:GetNames()

			nFilFlu := aScan(aDados,{|x| x == "filial"})
			nConta 	:= aScan(aDados,{|x| x == "conta"})
			nData 	:= aScan(aDados,{|x| x == "data"})
			nCCusto := aScan(aDados,{|x| x == "ccusto"})

			If nFilFlu = 0
				cError += "Erro estrutura. Falta campo filial"
				lRet := .F.
			EndIf
			If nConta = 0
				cError += "Erro estrutura. Falta campo conta"
				lRet := .F.
			EndIf
			If nData = 0
				cError += "Erro estrutura. Falta campo data"
				lRet := .F.
			EndIf
			If nCCusto = 0
				cError += "Erro estrutura. Falta campo ccusto"
				lRet := .F.
			EndIf

			If lRet //Passou na validação da estrutura
				cFilFlu := AllTrim(oJson:GetJsonText(aDados[nFilFlu]))
				cConta 	:= oJson:GetJsonText(aDados[nConta])
				cData 	:= oJson:GetJsonText(aDados[nData])
				cCCusto := oJson:GetJsonText(aDados[ncCusto])

				If Empty(cFilFlu)
					cError += "Campo obrigatório não preenchido - Filial"
					lRet := .F.
				EndIf
				If Empty(cConta)
					cError += "Campo obrigatório não preenchido - Conta"
					lRet := .F.
				EndIf
				If Empty(cData)
					cError += "Campo obrigatório não preenchido - Data"
					lRet := .F.
				EndIf
				If Empty(cCCusto)
					cError += "Campo obrigatório não preenchido - Centro Custo"
					lRet := .F.
				EndIf
			EndIf

			If lRet
				aRet := fRetSaldo(cFilFlu,cConta,cData,cCCusto)

				cEmail := aRet[3]
				cError := aRet[2]
				nSaldo := aRet[1]

				// If !Empty(cError)
				If Empty(cError)

					If nSaldo > 0
						::SetResponse('{"SALDO":' + FWJsonSerialize(nSaldo,.F.,.F.) + ',"APROV":"'+cEmail+'"}')
					Else
						::SetResponse('{"SALDO":' + FWJsonSerialize(0,.F.,.F.) + ',"APROV":"'+cEmail+'"}')
					EndIf
				Else
					lRet := .F.
				EndIf
			EndIf
		ELSE
			cError := "Erro na estrutura do Json."
			lRet = .F.
		EndIf

		If !lRet
			self:setStatus(400)
			SetRestFault(400,cError, .T.)
		EndIf

		ErrorBlock(oLastError)
	End SEQUENCE

	If !Empty(cErr)
		self:setStatus(400)
		SetRestFault(400,cErr, .T.)
		lRet := .F.
	EndIf


Return lRet

Static Function Aprova(oJson,lNovo)
	Local cDocumento      := ""       //-- Recebe o número do documento a ser avaliado
	Local cTipo     := ""       //-- Recebe o tipo do documento a ser avaliado
	Local cAprov    := ""       //-- Recebe o código do aprovador do documento
	Local lOk       := .T.      //-- Controle de validação e commit
	Local cCodFluig := ""
	Local cItGrp	:= ""
	Local cGrp	:= ""//NP3-09/07/2022
	//Local aErro     := {}       //-- Recebe msg de erro de processamento

	Default lNovo := .F.

	aDados := oJson:GetNames()
	nPosCNPJ := aScan(aDados,{|x| x == "cnpj"})
	nPosDoc  := aScan(aDados,{|x| x == "documento"})
	nPosApro := aScan(aDados,{|x| x == "aprovador"})
	nPosOk   := aScan(aDados,{|x| x == "aprovado"})
	nPosTipo := aScan(aDados,{|x| x == "tipo"})
	nPosObs  := aScan(aDados,{|x| x == "observacao"})
	nPosCod  := aScan(aDados,{|x| x == "codfluig"})
	nPosITG  := aScan(aDados,{|x| x == "itemgrupo"})
	nPosGrp  := aScan(aDados,{|x| x == "grupo"})//NP3-09/07/2022

	// cCnpj := IIF(nPosCNPJ>0,oJson:GetJsonText(aDados[nPosCNPJ]),'')
	cCnpj := ""
	cDocumento  := oJson:GetJsonText(aDados[nPosDoc])
	cAprov:= oJson:GetJsonText(aDados[nPosApro])
	cOk   := oJson:GetJsonText(aDados[nPosOk])
	cTipo := oJson:GetJsonText(aDados[nPosTipo])
	cObs  := oJson:GetJsonText(aDados[nPosObs])
	cItGrp:= oJson:GetJsonText(aDados[nPosITG])
	cGrp  := oJson:GetJsonText(aDados[nPosGrp])//NP3-09/07/2022
	If nPosCod > 0
		cCodFluig := oJson:GetJsonText(aDados[nPosCod])
	Endif

	//Procura Filial para conectar
	//U_WSSeekFil(cCnpj)
	//-- Códigos de operações possíveis (vaariável cOK) :
	//--    "001" // Liberado
	//--    "002" // Estornar
	//--    "003" // Superior
	//--    "004" // Transferir Superior
	//--    "005" // Rejeitado
	//--    "006" // Bloqueio
	//--    "007" // Visualizacao

	// lOk := U_AprovPCA(cDocumento, cTipo, cAprov, cOk, cObs, cCnpj, cCodFluig, lNovo, cItGrp)//NP3-09/07/2022
	lOk := U_AprovPCA(cDocumento, cTipo, cAprov, cOk, cObs, cCnpj, cCodFluig, lNovo, cItGrp, cGrp)


Return lOk

User Function AprovPCA(cDocumento, cTp, cAprv, cCodApv, cObs, cCnpj, cCodFluig, lNovo, cItGrp,cGrp)

	Local oModel094 := Nil     //-- Objeto que receberá o modelo da MATA094
	Local cDoc      := cDocumento //-- Recebe o número do documento a ser avaliado
	Local cTipo     := cTp     //-- Recebe o tipo do documento a ser avaliado
	Local cAprov    := cAprv   //-- Recebe o código do aprovador do documento
	Local nLenSCR   := 0       //-- Controle de tamanho de campo do documento
	Local lOk       := .T.     //-- Controle de validação e commit
	Local aErro     := {}      //-- Recebe msg de erro de processamento
	Private lMsErroAuto 	 := .F.
	Private lMsHelpAuto 	 := .T.
	Private lAutoErrNoFile   := .T.

	If !lNovo
		nLenSCR := TamSX3("CR_NUM")[1] //-- Obtem tamanho do campo CR_NUM
		DbSelectArea("SCR")
		// SCR->(DbSetOrder(4)) //-- CR_FILIAL+CR_TIPO+CR_NUM+CR_APROV
		SCR->(DBORDERNICKNAME( "ITGRP" )) //CR_FILIAL+CR_TIPO+CR_NUM+CR_ITGRP+CR_APROV

		// If SCR->(DbSeek(xFilial("SCR") + Padr(cTipo, len(SCR->CR_TIPO)) + Padr(cDoc, nLenSCR) + Padr(cItGrp, len(SCR->CR_ITGRP)) + Padr(cAprov, len(SCR->CR_APROV)) ))//NP3-09/07/2022
		If SCR->(DbSeek(xFilial("SCR") + Padr(cTipo, len(SCR->CR_TIPO)) + Padr(cDoc, len(SCR->CR_NUM)) + Padr(cItGrp, len(SCR->CR_ITGRP)) + Padr(cGrp, len(SCR->CR_GRUPO)) + Padr(cAprov, len(SCR->CR_APROV)) ))
			If SCR->CR_STATUS == '02'
				__cUserID  := SCR->CR_USER //"000199"

				//-- Códigos de operações possíveis:
				//--    "001" // Liberado
				//--    "002" // Estornar
				//--    "003" // Superior
				//--    "004" // Transferir Superior
				//--    "005" // Rejeitado
				//--    "006" // Bloqueio
				//--    "007" // Visualizacao

				//-- Seleciona a operação de aprovação de documentos
				A094SetOp(cCodApv)

				//-- Carrega o modelo de dados e seleciona a operação de aprovação (UPDATE)
				oModel094 := FWLoadModel('MATA094')
				oModel094:SetOperation( MODEL_OPERATION_UPDATE )
				oModel094:Activate()

				IF cCOdApv == "005"
					//-- Preenche justificativa
					oModel094:GetModel('FieldSCR'):SetValue('CR_OBS', cObs)
					// oModel094:Refresh()
				EndIf

				//-- Valida o formulário
				lOk := oModel094:VldData()

				If lOk
					//-- Se validou, grava o formulário
					lOk := oModel094:CommitData()
					Sleep(1000)
				EndIf

				//-- Avalia erros
				If !lOk
					//-- Busca o Erro do Modelo de Dados
					aErro := oModel094:GetErrorMessage()
					cMensagem := AllToChar(aErro[06])
					ConOut("-----------> NP3 ----- " + "Erro na aprovação: " + cMensagem)
				Else

					//-- Desativa o modelo de dados
					oModel094:DeActivate()

					__cUserID := ""
					If cCodApv == "001"
						// U_APROVAWF(cPedido,3)
						ConOut("-----------> NP3 ----- " + "Aprovado e enviando cancelamento dos mesmos níveis.")
						U_APROVAWF(cDocumento,5,cTipo,cItGrp,.T.,cGrp)//NP3-09/07/2022
					ElseIf cCodApv == "005"
						ConOut("-----------> NP3 ----- " + "Rejeitado e enviando cancelamento dos mesmos níveis.")
						U_APROVAWF(cDocumento,5,cTipo,cItGrp,.F.,cGrp,cCodFluig)
					Endif
				EndIf
			EndIf
		Else
			lOk := .F.
			cMensagem := "Documento nao encontrado"
			ConOut("-----------> NP3 ----- " + cMensagem)
		EndIf
	Else
		//Envia o Cancelamento antes para enviar a aprovação
		ConOut("-----------> NP3 ----- " + "Verificar se tem novos aprovadores")
		// U_CancPed(cDocumento,cCnpj,cCodfluig)

		U_APROVAWF(cDocumento,3,cTipo,cItGrp,,cGrp)
	EndIf

Return lOk



Static Function fRetSaldo(cFilFlu,cConta,cData,cCCusto)
	Local aRet   := {}
	Local cSql   := ""
	Local cZ20   := GetNextAlias()
	Local cORC   := GetNextAlias()
	Local dDt    := CTOD(cData)
	Local nMes   := MONTH(dDt)
	Local cMes   := StrZero(MONTH(dDt),2)
	Local cAno   := cValToChar(YEAR(dDt))
	Local nI     := 0
	Local cError := ""
	Local cFrom  := ""
	Local cInnerCta := ""
	Local cInnerCC := ""
	Local cWhere := ""
	Local nSaldo := 0
	Local cEmail := ""

	Local oLastError := ErrorBlock( { |e| cError := e:ErrorStack, Break(e) } )

	Z20->(dbSetOrder(1)) //Z20_FILIAL+Z20_CODIGO+Z20_ANO
	Z21->(dbSetOrder(1)) //Z21_FILIAL+Z21_CODIGO+Z21_CONTA


	BEGIN SEQUENCE


		cSql := " SELECT "
		cSql += " 	Z20_PRIORI "
		cSql += "  ,Z20_VALIDA "
		cSql += "  ,Z20_APROV "
		CSql += "  ,Z20_CODIGO"
		CSql += "  ,Z20_REVISA"
		CSql += "  ,Z20_ANO"
		cSql += " FROM "
		cSql += RetSqlName("Z20")+ " Z20 "
		cSql += " INNER JOIN " + RetSqlName("Z21") + " Z21 ON Z21_FILIAL = Z20_FILIAL "
		cSql += " 	AND Z21_CODIGO = Z20_CODIGO "
		cSql += " 	AND Z21_CONTA = '"+cConta+"' "
//	cSql += " 	AND Z21_CCUSTO = '"+cCCusto+"' "
		cSql += " 	AND Z21.D_E_L_E_T_='' "
		cSql += " WHERE Z20.D_E_L_E_T_='' "
		cSql += " 	AND Z20_ANO = '"+cAno+"'"
		cSql += "   AND Z20_STATUS = 'L'"

		MpSysOpenQuery(cSql,(cZ20))

		IF !(cZ20)->(Eof())

			//Busca e-mail do aprovador
			cEmail := Alltrim(UsrRetMail((cZ20)->Z20_APROV))

			cPrioriza := (cZ20)->Z20_PRIORI
			cValida   := (cZ20)->Z20_VALIDA
			cCodigo   := AllTrim((cZ20)->Z20_CODIGO)

			// cSql := " SELECT "
			// cSql += " 	Z20_PRIORI "
			// cSql += "  ,Z20_VALIDA "
			// cSql += "  ,Z20_APROV "
			// cSql += "  ,Z21_CONTA "
			// cSql += "  ,Z21_CCUSTO "
			// cSql += "  ,Z21_SMES01,Z21_SMES02,Z21_SMES03,Z21_SMES04,Z21_SMES05,Z21_SMES06 "
			// cSql += "  ,Z21_SMES07,Z21_SMES08,Z21_SMES09,Z21_SMES10,Z21_SMES11,Z21_SMES12 "

			cFrom := " FROM "
			cFrom += RetSqlName("Z20")+ " Z20 "

			cInnerCta := " INNER JOIN " + RetSqlName("Z21") + " Z21 ON Z21_FILIAL = Z20_FILIAL AND Z21_CODIGO = Z20_CODIGO AND Z20_REVISA = Z21_REVIS AND Z20_ANO = Z21_ANO AND Z21_CONTA = '"+cConta+"' AND Z21.D_E_L_E_T_='' "
			cInnerCC  := " INNER JOIN " + RetSqlName("Z21") + " Z21 ON Z21_FILIAL = Z20_FILIAL AND Z21_CODIGO = Z20_CODIGO AND Z20_REVISA = Z21_REVIS AND Z20_ANO = Z21_ANO AND Z21_CCUSTO = '"+cCCusto+"' AND Z21.D_E_L_E_T_='' "

			cWhere := " WHERE Z20.D_E_L_E_T_='' "
			cWhere += " AND Z20_CODIGO = '"+cCodigo+"'"

			cSql := " SELECT "

			DO CASE
			CASE cPrioriza == "1"//Grupo
				If cValida == '1'  //Mês
					cSql += 'Z20_SMES'+cMes + " AS SALDO "
				Else
					For nI := 1 to (IIf(cValida='2',12,nMes))
						cSql += Iif(nI>1,'+',' ')+"Z20_SMES"+StrZero(nI/*nMes*/,2)
					Next
					cSql := cSql + " AS SALDO "
				EndIf

				cSql += cFrom + cWhere

			CASE cPrioriza = '2' //Conta
				cSql += " SUM("
				If cValida == '1'  //Mês
					cSql += ' Z21_SMES'+cMes
				Else
					For nI := 1 to (IIf(cValida='2',12,nMes))
						cSql += Iif(nI>1,'+',' ')+"Z21_SMES"+StrZero(nI/*nMes*/,2)
					Next
				EndIf

				cSql += ") AS SALDO "
				cSql += cFrom + cInnerCta + cWhere
				cSql += " GROUP BY Z21_CONTA "

			CASE cPrioriza = '3' //CCusto
				cSql += " SUM("
				If cValida == '1'  //Mês
					cSql += ' Z21_SMES'+cMes
				Else
					For nI := 1 to (IIf(cValida='2',12,nMes))
						cSql += Iif(nI>1,'+',' ')+"Z21_SMES"+StrZero(nI/*nMes*/,2)
					Next
				EndIf

				cSql += ") AS SALDO "
				cSql += cFrom + cInnerCC + cWhere
				cSql += " GROUP BY Z21_CCUSTO "

			CASE cPrioriza = '4' //CCusto + Conta
				cSql += " SUM("
				If cValida == '1'  //Mês
					cSql += ' Z21_SMES'+cMes
				Else
					For nI := 1 to (IIf(cValida='2',12,nMes))
						cSql += Iif(nI>1,'+',' ')+"Z21_SMES"+StrZero(nI/*nMes*/,2)
					Next
				EndIf

				cSql += ") AS SALDO "
				cSql += cFrom + cInnerCC + cWhere + " AND Z21_CONTA = '"+cConta+"'"
				cSql += " GROUP BY Z21_CONTA,Z21_CCUSTO "

			END CASE

			MemoWrite("WSRETSALDO.SQL",cSql)


			MpSysOpenQuery (cSql,(cORC))

			If !(cORC)->(Eof())
				nSaldo := (cORC)->SALDO
			Else
				nSaldo := 0
				//cError := "Orçamento não encontrado para o grupo "+(cZ20)->Z20_DESC
			EndIf

			(cORC)->(dbCloseArea())
		else
			nSaldo := 0
			//cError := "Conta + Ano não encontrada : "+cConta + "/"+cAno
		EndIf

		(cZ20)->(dbCloseArea())

		ErrorBlock(oLastError)

	END SEQUENCE

	aAdd(aRet, nSaldo)
	aAdd(aRet, cError)
	aAdd(aRet, cEmail)

Return aRet

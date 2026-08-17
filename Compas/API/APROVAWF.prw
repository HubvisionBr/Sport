#Include "TBICONN.ch"
#include 'TOTVS.ch'

User Function APROVAWF(cDocumento,nOpc,cTipo,cGrpIT,lAprov,cGrupo,cCodFluig)

	Local aAprova := {}
	Local cCnpj
	Local cRet
	Local cCpdFlu := ""
	Local cAliasNP3 := GetNextAlias()

	Default nOpc := 0
	Default lAprov := .T.
	Default cGrupo := ""
	Default cCodFluig := ""

	If nOpc == 0
		Return
	EndIf


	If nOpc == 5
		If lAprov

			IF Select(cAliasNP3) > 0
				(cAliasNP3)->(DbCloseArea())
			Endif

			BEGINSQL ALIAS cAliasNP3

			SELECT CR_XIDFLU FROM %table:SCR% A
			WHERE A.%notDel%
			AND CR_FILIAL = %XFilial:SCR%
			AND CR_TIPO   = %Exp:cTipo%
			AND CR_NUM    = %Exp:cDocumento%
			AND CR_GRUPO  = %Exp:cGrupo%
			AND CR_ITGRP  = %Exp:cGrpIT%
			AND (CR_STATUS  = '05' OR CR_STATUS = '07')

			ENDSQL

			IF !(cAliasNP3)->(Eof())

				While !(cAliasNP3)->(Eof())

					cCpdFlu := AllTrim((cAliasNP3)->CR_XIDFLU)
					ConOut("-----------> NP3 ----- " + "Cancenlando o ID " + cCpdFlu + " no FLUIG.")

					U_CancPed(cDocumento,cCnpj,cCpdFlu,lAprov)

					(cAliasNP3)->(DbSkip())
				EndDo
			Endif
		Else
			ConOut("-----------> NP3 ----- " + "Cancelando o atividades APROVWF.")
			U_CancPed(cDocumento,cCnpj,cCodFluig,,cTipo)
		EndIf
	Else
		//Envia o cancelamento s  na altera  o
		IF nOpc == 4
			//Envia os cancelamentos
			cRet := U_CancPed(cDocumento,cCnpj,,,cTipo)
			IF !Empty(cRet)
				IF ValType(cRet) == "C"
					IF !"SUCCESS" $ Upper(cRet)
						Return
					Endif
				Else
					ConOut("Nao conseguiu enviar o cancelamento - provavelmente fora do ar o fluig")
				EndIf
			Endif


		Endif

		aAprova := fnAprovPC(cDocumento,cTipo,cGrpIT,cGrupo)

		IF Len(aAprova) > 0
			MailAprov(aAprova,cDocumento,cTipo)
		Else
			//ConOut("N o foi gerada aprova  o para o pedido.")
			ConOut("Nao foi gerada aprovacao para o pedido " +  iif(cDocumento == NIL, "Nil",cDocumento))
		Endif
	EndIf

Return


/*---------------------------------------------------------------------------
--  Fun  o: Constroi vetor com dados dos aprovadores do Ped. de compra.    --
--          Passado como parametros.                                       --
--           Parametros:                                                   --
--             Recebe cNumero - Codigo do Pedido para Libera  o.           --
--                          f                                               --
--             Devolve aAprovador - Vetor com os dados dos Aprovedores     --
--                                  do Pedido de Liberacao Descritos em    --
--                                  SCR - Doc's por alcada, sendo:         --
--                        aAprovador[01] - Codigo do Grupo de Aprovadores  --
--                        aAprovador[02] - Codigo do Aprovador             --
--                        aAprovador[03] - Codigo do usuario correspondete --
--                        aAprovador[04] - Nome                            --
--                        aAprovador[05] - Endereco de e-mail              --
--                        aAprovador[06] - Tipo de Aprovacao (Liberacao    --
--                                         ou Visto)                       --
--                        aAprovador[07] - "S" ou "N" Considera Limites    --
--                        aAprovador[08] - Tipo de Liberacao:              --
--                                         "U" - Usuario - Libera apenas   --
--                                               seu usuario               --
--                                         "N" - Pode Liberar todo o nivel --
--                                               a que este pertence       --
--                                         "P" - Libera todo o documento,  --
--                                               independente de outras    --
--                                               aprovacoes (autonomia     --
--                                               total.                    --
--                        aAprovador[09] - Tipo 			               --
-----------------------------------------------------------------------------*/
Static Function fnAprovPc(cNumero,cTipo,cGrpIT,cGrupo)
	Local aAprovador:= {}
	Local lApvPC	:= .F.
	Local cAprov	:= ""
	Local nNivel    := 1
	Local nCount    := 1
	// Local cUser_Id  := __cuserid 

	ConOut("########### PARAMETROS PARA BUSCA DE DOCUMENTOS DA SCR ########")
	conOut("cNumero:" + iif(cNumero == NIL, "Nil",cNumero))
	conOut("cTipo:" + iif(cTipo == NIL, "Nil",cTipo))
	conOut("cGrpIT:" + iif(cGrpIT == NIL, "Nil",cGrpIT))
	conOut("cGrupo:" + iif(cGrupo == NIL, "Nil",cGrupo))

	dbSelectArea("SCR")
	// SCR->(dbSetorder(1))
	SCR->(DBORDERNICKNAME( "ITGRP" )) //CR_FILIAL+CR_TIPO+CR_NUM+CR_ITGRP 
	// IF SCR->(dbSeek(xFilial("SCR") + PADR(cTipo,LEN(SCR->CR_TIPO)) + PADR(cNumero,LEN(SCR->CR_NUM)) + PADR(cGrpIT,LEN(SCR->CR_ITGRP)) ))//NP3-09/07/2022
	IF SCR->(dbSeek(xFilial("SCR") + PADR(cTipo,LEN(SCR->CR_TIPO)) + PADR(cNumero,LEN(SCR->CR_NUM)) + PADR(cGrpIT,LEN(SCR->CR_ITGRP))  + PADR(cGrupo,LEN(SCR->CR_GRUPO)) ))

		/*
		{ 'CR_STATUS== "01"', 'BR_AZUL' },;//Bloqueado p/ sistema(aguardando outros niveis)
		{ 'CR_STATUS== "02"', 'DISABLE' },;//Aguardando Liberacao do usuario
		{ 'CR_STATUS== "03"', 'ENABLE'  },;//Pedido Liberado pelo usuario
		{ 'CR_STATUS== "04"', 'BR_PRETO'},;//Pedido Bloqueado pelo usuario
		{ 'CR_STATUS== "05"', 'BR_CINZA'} }//Pedido Liberado por outro usuario
		*/

		// ---- Loop em documentacao p/alcada para verificar quem deve aprovar a liberacao
		// While !(SCR->(Eof())) .and. SCR->CR_FILIAL == xFilial("SCR") .and. SCR->CR_TIPO == cTipo .and. Alltrim(SCR->CR_NUM) == cNumero .And. AllTrim(cGrpIT) == SCR->CR_ITGRP //NP3-09/07/2022
		While !(SCR->(Eof())) .and. SCR->CR_FILIAL == xFilial("SCR") .and. SCR->CR_TIPO == cTipo .and. Alltrim(SCR->CR_NUM) == cNumero .And. AllTrim(cGrpIT) == AllTrim(SCR->CR_ITGRP);
				.And. AllTrim(cGrupo) == SCR->CR_GRUPO
			lApvPC := .T.

			//-- Localiza O Aprovador e o Grupo de Aprovacao
			dbSelectArea("SAL")
			SAL->(dbSetorder(3))
			SAL->(dbSeek(xFilial("SAL") + SCR->CR_GRUPO + SCR->CR_APROV))

			//Verifica se o aprovador est  ausente
			dbSelectArea("SAK")
			SAK->(dbSetOrder(1))
			SAK->(MsSeek(xFilial("SAK")+SCR->CR_APROV))

			//Informa que ir  procurar o usu rio pelo ID
				/*
				1 - ID do usu rio/grupo
				2 - Nome do usu rio/grupo;
				3 - Senha do usu rio
				4 - E-mail do usu rio
				*/
			PswOrder(1)
			IF PswSeek(SCR->CR_USER,.t.)
				// __cUserID := SCR->CR_USER
			Endif

			//-- Verifica se esta aguardando liberacao e monta o ventor com os aprovadores do Grupo
			If Val(SCR->CR_STATUS) == 2 .and. PswSeek(SCR->CR_USER,.t.) .and. cAprov <> SCR->CR_APROV
				//Controle para saber qual n vel de aprova  o ser  enviado para o fluig
				IF nCount == 1
					nNivel :=  VAL(SCR->CR_NIVEL)
					nCount++
				Endif

				IF VAL(SCR->CR_NIVEL) == nNivel
					nNivel := VAL(SCR->CR_NIVEL)
					cAprov := SCR->CR_APROV
					//aInfo := PswRet(1) //http://tdn.totvs.com/pages/releaseview.action?pageId=267792734

					//-- Monta vetor dos aprovadores {[Grupo de Apr.],[Aprovador],[USuario],[Nome],[e-mail],[Tipo de Aprovacao],[Considera Limites],[Tipo Lib.]}
					aAdd(aAprovador, {SCR->CR_GRUPO,;
						SCR->CR_APROV,;
						SCR->CR_USER,;
						FwGetUserName(SCR->CR_USER),;
						AllTrim(UsrRetMail(SCR->CR_USER)),;
						SAL->AL_LIBAPR,;
						SAL->AL_AUTOLIM,;
						SAL->AL_TPLIBER,;
						cTipo,;
						SCR->CR_ITGRP})
				Endif
			Endif

			SCR->(dbSkip())
		Enddo
	Else
	 //ConOut("### NAO ENCONTROU DOCUMENTO " + cNumero + " NA SCR ######")
	Endif
// __cuserid := cUser_Id
Return aAprovador


/*--------------------------------------------------
--  Fun  o: Envia e-mail para os aprovadores.     --
--                                                --
----------------------------------------------------*/
Static Function MailAprov(aAprovPc,cDocumento,cTipo)

	Local oJson
	Local oPedido
	//Local oItens
	Local oForms
	Local nX //,nZ,nY
	Local aItens := {}
	Local cAprov,cGrp,cItGrp,cApr := ""

	For nX := 1 to Len(aAprovPc)

		//Verifica se o pedido j  foi enviado para o Fluig

		aItens := {}
		aAprovacao := {}
		cAprov := aAprovPc[nX,3] //USER
		cApr := aAprovPc[nX,2] //Aprovador
		cGrp := aAprovPc[nX,1] //GRUPO
		cItGrp := aAprovPc[nX,10] //ITEM DO GRUPO
		aAdd(aAprovacao,aAprovPc[nX])
		oPedido := JsonObject():New()
		oPedido['targetState' ] := 35
		oPedido['targetAssignee' ] := "admin"
		oPedido['subProcessTargetState' ] := 0
		oPedido['comment' ] := "Solicita  o iniciada automaticamente"
		oForms := JsonObject():new()

		conOut(" HV -----> Criando formulário AddForms")

		AddForms(aAprovacao,@oForms,cDocumento,@aItens,cTipo)

		IF oForms <> Nil
			conOut(" HV -----> Formilário contém itens ")

			oPedido['formFields'] := oForms

			conOut(" HV -----> Enviando para o Fluig")
			cRet := U_ConFluig(oPedido:ToJson(),'/process-management/api/v2/processes/liberacaoDocumento/start')

			Sleep(500)

			IF !Empty(cRet)
				FwJSONDeserialize(cRet, @oJSON)
				If ValType(ojson) == "O"
					IF ValType(ojson:processinstanceid) == "N"
						//alteracao para incluir o id de aprovacao no pedido de compras RODRIGO SOBRAL 2024/04/24
						If RecLock("SC7", .F.)
							SC7->C7_XIDFLUI:=AllTrim(STR(ojson:processinstanceid))
							SC7->(MsUnlock())
						EndIf			
						SC7->(DbSkip())

							DbSelectArea("DBM")
							DBM->(DbSetOrder(1))
							IF DBM->(MsSeek(xFilial("DBM") + PADR(cTipo,LEN(DBM->DBM_TIPO)) + padr(cDocumento,LEN(DBM->DBM_NUM)) + padr(cGrp,LEN(DBM->DBM_GRUPO)) +;
									padr(cItGrp,LEN(DBM->DBM_ITGRP)) + padr(cAprov,LEN(DBM->DBM_USER))))

								While DBM->DBM_FILIAL == xFilial("DBM") .AND. DBM->DBM_TIPO == PADR(cTipo,LEN(DBM->DBM_TIPO)) .and. DBM->DBM_NUM == padr(cDocumento,LEN(DBM->DBM_NUM));
										.and. DBM->DBM_GRUPO == padr(cGrp,LEN(DBM->DBM_GRUPO)) .and. DBM->DBM_ITGRP == padr(cItGrp,LEN(DBM->DBM_ITGRP)) .and. DBM->DBM_USER == padr(cAprov,LEN(DBM->DBM_USER)) 
									
									If RecLock("DBM", .F.)
										DBM->DBM_XIDFLU := AllTrim(STR(ojson:processinstanceid))
										DBM->(MsUnlock())
									EndIf
									
									DBM->(DbSkip())
								EndDo

							ENDIF
							DbSelectArea("SCR")
							SCR->(DBORDERNICKNAME( "ITGRP" )) //CR_FILIAL+CR_TIPO+CR_NUM+CR_ITGRP 
							IF SCR->(MsSeek(xFilial("SCR") + PADR(cTipo,LEN(SCR->CR_TIPO)) + padr(cDocumento,LEN(SCR->CR_NUM)) + padr(cItGrp,LEN(SCR->CR_ITGRP)) ;
								+ padr(cGrp,LEN(SCR->CR_GRUPO)) + PADR(cApr,LEN(SCR->CR_APROV)) ))
									
								While xFilial("SCR") == SCR->CR_FILIAL .and. PADR(cTipo,LEN(SCR->CR_TIPO)) == SCR->CR_TIPO .and. padr(cDocumento,LEN(SCR->CR_NUM)) == SCR->CR_NUM .and. padr(cItGrp,LEN(SCR->CR_ITGRP)) == SCR->CR_ITGRP;
									.and. padr(cGrp,LEN(SCR->CR_GRUPO)) == SCR->CR_GRUPO .and. PADR(cApr,LEN(SCR->CR_APROV)) == SCR->CR_APROV
									
									If RecLock("SCR", .F.)
										SCR->CR_XIDFLU := AllTrim(STR(ojson:processinstanceid))
										SCR->(MsUnlock())
									EndIf
									
									SCR->(DbSkip())
								EndDo
							EndIf	
						FreeObj(oJson)
					Endif
				else
					
				EndIf
			EndIf
		Else
			conOut(" HV -----> Sem registro de itens no Forms")
			oForms := JsonObject():new()
		EndIf
		FreeObj(oForms)
		FreeObj(oPedido)
	Next nX

Return

Static Function AddForms(aAprovPc,oForms,cDocumento,aItens,cTipo)

	Local i
	Local cNmUlAprv := ""
	Local cDtUlAprv := ""
	Local cObjt := ""
	Local cFornec := ""
	Local cLoja := ""
	Local dEmissao := STOD(" ")
	Local cContra := ""
	Local cMed := ""
	Local cRev := ""
	Local cCCusto := ""
	Local cProd := ""
	Local cDesc := ""
	Local cQuant := ""
	Local cPrc 	:= ""
	Local cTot 	:= ""
	Local cObs := ""
	Local cGrpAprov	:= ""
	Local cItGrp	:= ""
	Local cITMed := ""
	Local aForm := {}
	Local cPlan := ""
	Local cClvl := ""
	Local cItemC := ""
	Local cConta := ""
	Local cCCDesc := ""
	Local cCT1Desc := ""
	Local cCTDDesc := ""
	Local cCTHDesc := ""
	Local cSolicit := ""
	Local cNumSc   := ""
	Local cComprador := ""
	Local cUserComp := ""
	Local cEmailComp := ""
	Local nTot	:= 0

	cPlan := SubStr(cDocumento, 7)
	cPlan := SubStr(cPlan, 1, Len(cPlan)-3)

	// DbSelectArea("SC7")
	// SC7->(DbSetOrder(1))
	// SC7->(MsSeek(xFilial("SC7") + cDocumento))

	// ---- Verificar  ltimo aprovador
	cQuery := "Select SCR.CR_USER, SCR.CR_DATALIB from " + RetSqlName("SCR") + " SCR"
	cQuery += "   where SCR.D_E_L_E_T_ <> '*'"
	cQuery += "     and SCR.CR_FILIAL = '" + xFilial("SCR") + "'"
	cQuery += "     and SCR.CR_STATUS = '03'"
	cQuery += "     and SCR.CR_NUM    = '" + cDocumento + "'"
	cQuery += "     and SCR.CR_TIPO   = '" + cTipo +"' "

	cQuery += "  Order by SCR.CR_FILIAL, SCR.CR_DATALIB desc"
	cQuery := ChangeQuery(cQuery)
	dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"QAPV",.F.,.T.)

	If ! QAPV->(Eof())
		cNmUlAprv := UsrFullName(QAPV->CR_USER)
		cDtUlAprv := DToC(SToD(QAPV->CR_DATALIB))
	EndIf

	QAPV->(dbCloseArea())
	// -------------------------------

	//Enviar um e-mail para cada aprocador do mesmo n vel
	For i := 1 to Len(aAprovPc)

		DbSelectArea("DBM")
		DBM->(DbSetOrder(1))
		conOut(" HV -----> Procurando na DBM o documento ")
		IF DBM->(MsSeek(xFilial("DBM") + aAprovPc[i][9] + cDocumento))
			cQuery    := " SELECT DISTINCT SCR.CR_NUM,DBM.DBM_ITEM,DBM.DBM_ITEMRA "
			cQuery	  += " FROM "+RetSqlName("SCR")+" SCR LEFT JOIN "
			cQuery	  += RetSqlName("DBM")+" DBM ON "
			cQuery	  += " CR_TIPO=DBM_TIPO AND "
			cQuery	  += " CR_NUM=DBM_NUM AND "
			cQuery	  += " CR_GRUPO=DBM_GRUPO AND "
			cQuery	  += " CR_ITGRP=DBM_ITGRP AND "
			cQuery	  += " CR_USER=DBM_USER AND "
			cQuery	  += " CR_USERORI=DBM_USEROR AND "
			cQuery	  += " CR_APROV=DBM_USAPRO AND "
			cQuery	  += " CR_APRORI=DBM_USAPOR AND "
			cQuery    += " DBM.D_E_L_E_T_<> '*' AND "
			cQuery    += " DBM.DBM_XIDFLU = '' AND "
			cQuery    += " DBM.DBM_APROV = '2' "

			cQuery    += " WHERE SCR.CR_FILIAL='"+xFilial("SCR")+"' AND "
			cQuery    += " SCR.D_E_L_E_T_  <> '*' AND "
			cQuery    += " SCR.CR_APROV = '"+aAprovPc[i][2]+"' AND "
			cQuery    += " SCR.CR_NUM = '"+Padr(cDocumento,Len(SCR->CR_NUM))+"' AND "

			cQuery    += " SCR.CR_TIPO = '" + cTipo +"' AND"
			cQuery    += " SCR.CR_ITGRP = '" + aAprovPc[i][10] + "' AND"
			cQuery    += " SCR.CR_GRUPO = '" + aAprovPc[i][1] + "' "//NP3-09/07/2022

			conOut(" HV -----> Query: " + cQuery)

			IF Select("QTMP") > 0
				QTMP->(DbCloseArea())
			EndIf

			dbUseArea(.T.,"TOPCONN",TcGenQry(,,cQuery),"QTMP",.F.,.T.)

			nCount := 1

			While ! QTMP->(Eof())
				If cTipo == 'IP'
					DbSelectArea("SC7")
					SC7->(DbSetOrder(1))
					SC7->(DbSeek(xFilial("SC7") + SubStr(QTMP->CR_NUM,1,TamSx3("C7_NUM")[1]) + QTMP->DBM_ITEM))

					IF Alltrim(SC7->C7_TPFRETE) == "C"
						cFrete := "CIF"
					ElseIF  Alltrim(SC7->C7_TPFRETE) == "F"
						cFrete := "FOB"
					Else
						cFrete := "Nao informado"
					Endif
					
					IF !Empty(SC7->C7_CONTRA)
						DbSelectArea("CN9")
						CN9->(DbSetOrder(1))
						CN9->(MsSeek(xFilial("CN9") + SC7->C7_CONTRA + SC7->C7_CONTREV))
						cObjt := MsMM(CN9->CN9_CODOBJ)
					EndIf

					cFornec 	:= ALLTRIM(SC7->C7_FORNECE)
					cLoja 		:= ALLTRIM(SC7->C7_LOJA)
					dEmissao 	:= DToc(SC7->C7_EMISSAO)
					cContra 	:= ALLTRIM(SC7->C7_CONTRA)
					cMed 		:= ALLTRIM(SC7->C7_MEDICAO)
					cRev 		:= ALLTRIM(SC7->C7_CONTREV)
					cCCusto 	:= ALLTRIM(SC7->C7_CC)
					cClvl		:= ALLTRIM(SC7->C7_CLVL)
					cItemC		:= ALLTRIM(SC7->C7_ITEMCTA)
					cConta		:= ALLTRIM(SC7->C7_CONTA)
					cProd 		:= ALLTRIM(SC7->C7_PRODUTO)
					cQuant 		:= Transform(SC7->C7_QUANT,PesqPict("SC7","C7_QUANT"))
					cPrc 		:= Transform(SC7->C7_PRECO,PesqPict("SC7","C7_PRECO"))
					nTot 		:= (SC7->C7_TOTAL+SC7->C7_SEGURO +SC7->C7_VALFRE+SC7->C7_VALIPI+SC7->C7_DESPESA)-SC7->C7_VLDESC 
					cTot := Transform(nTot,PesqPict("SC7","C7_TOTAL"))
					//RODRIGO SOBRAL 19/10/2023

					cObs 		:= Iif(!Empty(SC7->C7_OBS),SC7->C7_OBS,SC7->C7_OBSM)
					cGrpAprov	:= aAprovPc[i][1]
					cItGrp		:= aAprovPc[i][10]
					cITMed 		:=IIF(empty(SC7->C7_ITEMED),AllTrim(SC7->C7_ITEM),AllTrim(SC7->C7_ITEMED))
					cComprador  := FwGetUserName(SC7->C7_USER)//UsrRetName((cAliasSC7)->C7_USER)
    				cUserComp   := SC7->C7_USER
					cEmailComp	:= Alltrim(UsrRetMail(SC7->C7_USER))

					//NP3-02/08/2022
					DbSelectArea("SC1")
					SC1->(DbSetOrder(2))//C1_FILIAL+C1_PRODUTO+C1_NUM+C1_ITEM
					IF SC1->(DbSeek(xFilial("SC1") + Padr(SC7->C7_PRODUTO,Len(SC1->C1_PRODUTO)) + Padr(SC7->C7_NUMSC,Len(SC1->C1_NUM)) + Padr(SC7->C7_ITEMSC,Len(SC1->C1_ITEM))   ))
						//Adicionado o ID da SC e o Solicitante -- 15/08/2022
						cSolicit := SC1->C1_SOLICIT
						cNumSc   := SC1->C1_NUM
					Else
						//Adicionado o ID da SC e o Solicitante -- 15/08/2022
						cSolicit := ""
						cNumSc   := ""
					Endif
					
				elseiF cTipo == 'IM'
					DbSelectArea("CND")
					CND->(DbSetOrder(4))
					CND->(MsSeek(xFilial("CND") + PADR(SubStr(cDocumento,1,6),len(CND->CND_NUMMED))))

					DbSelectArea("CNE")
					CNE->(DbSetOrder(6))
					CNE->(MsSeek(xFilial("CNE") + PADR(SubStr(cDocumento,1,6),len(CNE->CNE_NUMMED)) + PADR(cPlan,LEN(CNE->CNE_NUMERO)) +;
					 PADR(QTMP->DBM_ITEM,LEN(CNE->CNE_ITEM))))      
					//  PADR(RIGHT(cDocumento,3),LEN(CNE->CNE_ITEM))))

					DbSelectArea("CXN")
					CXN->(DbSetOrder(1))
					CXN->(MsSeek(xFilial("CXN") + PADR(CND->CND_CONTRA,LEN(CXN->CXN_CONTRA)) + PADR(CND->CND_REVISA,len(CXN->CXN_REVISA)) +;
						PADR(SubStr(cDocumento,1,6),len(CXN->CXN_NUMMED)) +	PADR(cPlan,LEN(CXN->CXN_NUMPLA)) ))
					                          				

					DbSelectArea("CN9")
					CN9->(DbSetOrder(1))
					CN9->(MsSeek(xFilial("CN9") + CND->CND_CONTRA + CND->CND_REVISA))
					cObjt := MsMM(CN9->CN9_CODOBJ)				

					cFrete := ""

					cFornec 	:= ALLTRIM(CXN->CXN_FORNEC)
					cLoja 		:= ALLTRIM(CXN->CXN_LJFORN)
					dEmissao 	:= STOD("")
					cContra 	:= CNE->CNE_CONTRA//ALLTRIM(CNE->CNE_CONTRA )
					cMed		:= CNE->CNE_NUMMED//ALLTRIM(CNE->CNE_NUMMED)
					cRev		:= CNE->CNE_REVISA//ALLTRIM(CNE->CNE_REVISA)
					cCCusto		:= ALLTRIM(CNE->CNE_CC)
					cProd		:= ALLTRIM(CNE->CNE_PRODUT)
					cQuant		:= Transform(CNE->CNE_QUANT,PesqPict("CNE","CNE_QUANT"))
					cPrc 		:= Transform(CNE->CNE_VLUNIT,PesqPict("CNE","CNE_VLUNIT"))
					cTot 		:= Transform(CNE->CNE_VLTOT,PesqPict("CNE","CNE_VLTOT"))
					cObs		:= CNE->CNE_XOBSM//""
					cGrpAprov	:= aAprovPc[i][1]
					cItGrp		:= aAprovPc[i][10]
					cITMed		:= ALLTRIM(CNE->CNE_ITEM)
					cClvl		:= ALLTRIM(CNE->CNE_CLVL)
					cItemC		:= ALLTRIM(CNE->CNE_ITEMCT)
					cConta		:= ALLTRIM(CNE->CNE_CONTA)
					cComprador  := ""
    				cUserComp   := ""
					cEmailComp  := ""

					cSolicit := ""
					cNumSc   := ""

				EndIf

				IF nCount == 1

					cCdFilial  := FwCodFil()
					cCdEmpresa := FwCodEmp()
					cNmFilial  := FwFilialName()
					cNmEmpresa := FwEmpName(cCdEmpresa)

					DbSelectArea("SA2")
					SA2->(DbSetOrder(1))
					SA2->(MsSeek(xFilial("SA2") + Padr(cFornec,Len(SA2->A2_COD)) + Padr(cLoja,Len(SA2->A2_LOJA))))

					aForm := {}
					aAdd(aForm,cCdEmpresa)
					aAdd(aForm,cNmEmpresa)
					aAdd(aForm,cCdFilial)
					aAdd(aForm,cNmFilial)
					aAdd(aForm,AllTrim(cNmUlAprv))
					aAdd(aForm,cDtUlAprv)
					aAdd(aForm,aAprovPc[i][02])
					aAdd(aForm,cDocumento)
					aAdd(aForm,dEmissao)
					aAdd(aForm,cFrete)
					aAdd(aForm,"0")
					aAdd(aForm,"")
					aAdd(aForm,SA2->A2_COD + SA2->A2_LOJA + " - " + SA2->A2_NOME)
					aAdd(aForm,SA2->A2_END)
					aAdd(aForm,SA2->A2_TEL)
					aAdd(aForm,aAprovPc[i][9])
					aAdd(aForm,aAprovPc[i][5])
					aAdd(aForm,aAprovPc[i][04])
					aAdd(aForm,cContra)
					aAdd(aForm,cMed)
					aAdd(aForm,cRev)
					aAdd(aForm,cObjt)
					aAdd(aForm,cGrpAprov)
					aAdd(aForm,cItGrp)
					aAdd(aForm,cComprador)
					aAdd(aForm,cUserComp)
					aAdd(aForm,cEmailComp)
					MontForm(1,aForm,@oForms)
					
				Endif
				//Retorna descri  o do centro de custo
				cCCDesc  := Iif(Empty(cCCusto),'', Posicione("CTT",1,xFilial("CTT") + cCCusto,"CTT_DESC01"))
				cCC      := cCCusto + "-" + cCCDesc
				//Retorna descri  o da conta cont bil
				cCT1Desc := Iif(Empty(cConta),'', Posicione("CT1",1,xFilial("CT1") + cConta,"CT1_DESC01"))
				cCT1      := cConta + "-" + cCT1Desc
				//Retorna descri  o do item conta
				cCTDDesc := Iif(Empty(cItemC),'', Posicione("CTD",1,xFilial("CTD") + cItemC,"CTD_DESC01"))
				cCTD   := cItemC + "-" + cCTDDesc
				//Retorna descri  o da classe de valor
				cCTHDesc := Iif(Empty(cClvl),'', Posicione("CTH",1,xFilial("CTH") + cClvl,"CTH_DESC01"))
				cCTH    	 := cClvl + "-" + cCTHDesc

				DbSelectArea("SB1")
				SB1->(DbSetOrder(1))
				SB1->(MsSeek(xFilial("SB1") + Padr(cProd,Len(SB1->B1_COD)) ))

				cDesc := SB1->B1_DESC

				DbSelectArea("SB2")
				SB2->(DbSetOrder(1))
				SB2->(MsSeek(xFilial("SB2") +  Padr(cProd,Len(SB1->B1_COD))))
				
				aForm := {}

				aAdd(aForm,cITMed)
				aAdd(aForm,cProd)
				aAdd(aForm,cDesc)
				aAdd(aForm,cQuant)
				aAdd(aForm,cPrc)
				aAdd(aForm,cTot)//totalpedido
				aAdd(aForm,cCC)
				aAdd(aForm,DtoC(SB1->B1_UCOM))
				aAdd(aForm,Transform(SB1->B1_UPRC,PesqPict("SB1","B1_UPRC")))
				aAdd(aForm,cObs)
				aAdd(aForm,cITMed)
				aAdd(aForm,alltrim(cCTH))
				aAdd(aForm,alltrim(cCTD))
				aAdd(aForm,alltrim(cCT1))

				//Adicionado o ID da SC e o Solicitante -- 15/08/2022
				aAdd(aForm,cSolicit)
				aAdd(aForm,cNumSc)

				MontForm(2,aForm,@oForms)

				nCount++
				QTMP->(DbSkip())
			EndDo
		Else
			If  cTipo == 'PC'
				DbSelectArea("SC7")
				SC7->(DbSetOrder(1))
				SC7->(MsSeek(xFilial("SC7") + cDocumento))
				nCount := 1
				While !SC7->(Eof()) .and. cDocumento == SC7->C7_NUM
					IF nCount == 1

						cCdFilial  := FwCodFil()
						cCdEmpresa := FwCodEmp()
						cNmFilial  := FwFilialName()
						cNmEmpresa := FwEmpName(cCdEmpresa)

						cFornec 	:= ALLTRIM(SC7->C7_FORNECE)
						cLoja 		:= ALLTRIM(SC7->C7_LOJA)
						dEmissao 	:= DToC(SC7->C7_EMISSAO)
						cContra 	:= ALLTRIM(SC7->C7_CONTRA)
						cMed 		:= ALLTRIM(SC7->C7_MEDICAO)
						cRev 		:= ALLTRIM(SC7->C7_CONTREV)
						cCCusto 	:= ALLTRIM(SC7->C7_CC)
						cClvl		:= ALLTRIM(SC7->C7_CLVL)
						cItemC		:= ALLTRIM(SC7->C7_ITEMCTA)
						cConta		:= ALLTRIM(SC7->C7_CONTA)
						cProd 		:= ALLTRIM(SC7->C7_PRODUTO)
						cQuant 		:= Transform(SC7->C7_QUANT,PesqPict("SC7","C7_QUANT"))
						cPrc 		:= Transform(SC7->C7_PRECO,PesqPict("SC7","C7_PRECO"))
						//AJUSTE RODRIGO SOBRAL 23/01/2024
					    //adicao na formacao do valor total do produto 
					    nTot += (SC7->C7_TOTAL+SC7->C7_SEGURO +SC7->C7_VALFRE+SC7->C7_VALIPI+SC7->C7_DESPESA)-SC7->C7_VLDESC 
					    cTot:=Transform(nTot,PesqPict("SC7","C7_TOTAL"))
					    //RODRIGO SOBRAL 19/10/2023
						cObs 		:= Iif(!Empty(SC7->C7_OBS),SC7->C7_OBS,SC7->C7_OBSM)
						cGrpAprov	:= aAprovPc[i][1]
						cItGrp		:= aAprovPc[i][10]
						cITMed 		:=IIF(empty(SC7->C7_ITEMED),AllTrim(SC7->C7_ITEM),AllTrim(SC7->C7_ITEMED))
						cComprador  := FwGetUserName(SC7->C7_USER)//UsrRetName((cAliasSC7)->C7_USER)
    					cUserComp   := SC7->C7_USER
						cEmailComp	:= Alltrim(UsrRetMail(SC7->C7_USER))

						IF Alltrim(SC7->C7_TPFRETE) == "C"
							cFrete := "CIF"
						ElseIF  Alltrim(SC7->C7_TPFRETE) == "F"
							cFrete := "FOB"
						Else
							cFrete := "Nao informado"
						Endif

						DbSelectArea("SA2")
						SA2->(DbSetOrder(1))
						SA2->(MsSeek(xFilial("SA2") + Padr(cFornec,Len(SA2->A2_COD)) + Padr(cLoja,Len(SA2->A2_LOJA))))

						IF !Empty(SC7->C7_CONTRA)
							DbSelectArea("CN9")
							CN9->(DbSetOrder(1))
							CN9->(MsSeek(xFilial("CN9") + SC7->C7_CONTRA + SC7->C7_CONTREV))
							cObjt := MsMM(CN9->CN9_CODOBJ)
						EndIf

						aForm := {}
						aAdd(aForm,cCdEmpresa)
						aAdd(aForm,cNmEmpresa)
						aAdd(aForm,cCdFilial)
						aAdd(aForm,cNmFilial)
						aAdd(aForm,AllTrim(cNmUlAprv))
						aAdd(aForm,cDtUlAprv)
						aAdd(aForm,aAprovPc[i][02])
						aAdd(aForm,cDocumento)
						aAdd(aForm,dEmissao)
						aAdd(aForm,cFrete)
						aAdd(aForm,"0")
						aAdd(aForm,"")
						aAdd(aForm,SA2->A2_COD + SA2->A2_LOJA + " - " + SA2->A2_NOME)
						aAdd(aForm,SA2->A2_END)
						aAdd(aForm,SA2->A2_TEL)
						aAdd(aForm,aAprovPc[i][9])
						aAdd(aForm,aAprovPc[i][5])
						aAdd(aForm,aAprovPc[i][04])
						aAdd(aForm,cContra)
						aAdd(aForm,cMed)
						aAdd(aForm,cRev)
						aAdd(aForm,cObjt)
						aAdd(aForm,cGrpAprov)
						aAdd(aForm,cItGrp)
						aAdd(aForm,cComprador)
						aAdd(aForm,cUserComp)
						aAdd(aForm,cEmailComp)

						MontForm(1,aForm,@oForms)
						
					Endif

					//Retorna descri  o do centro de custo
					cCCDesc  := Iif(Empty(cCCusto),'', Posicione("CTT",1,xFilial("CTT") + cCCusto,"CTT_DESC01"))
					cCC      := cCCusto + "-" + cCCDesc
					//Retorna descri  o da conta cont bil
					cCT1Desc := Iif(Empty(cConta),'', Posicione("CT1",1,xFilial("CT1") + cConta,"CT1_DESC01"))
					cCT1      := cConta + "-" + cCT1Desc
					//Retorna descri  o do item conta
					cCTDDesc := Iif(Empty(cItemC),'', Posicione("CTD",1,xFilial("CTD") + cItemC,"CTD_DESC01"))
					cCTD   := cItemC + "-" + cCTDDesc
					//Retorna descri  o da classe de valor
					cCTHDesc := Iif(Empty(cClvl),'', Posicione("CTH",1,xFilial("CTH") + cClvl,"CTH_DESC01"))
					cCTH    	 := cClvl + "-" + cCTHDesc

					DbSelectArea("SB1")
					SB1->(DbSetOrder(1))
					SB1->(MsSeek(xFilial("SB1") + Padr(cProd,Len(SB1->B1_COD)) ))

					cDesc := SB1->B1_DESC

					DbSelectArea("SB2")
					SB2->(DbSetOrder(1))
					SB2->(MsSeek(xFilial("SB2") +  Padr(cProd,Len(SB1->B1_COD))))
					
					aForm := {}

					aAdd(aForm,StrZero(nCount,TamSX3("C7_ITEM")[1]))
					aAdd(aForm,cProd)
					aAdd(aForm,cDesc)
					aAdd(aForm,cQuant)
					aAdd(aForm,cPrc)
					aAdd(aForm,cTot)//totalpedido
					aAdd(aForm,cCC)
					aAdd(aForm,DtoC(SB1->B1_UCOM))
					aAdd(aForm,Transform(SB1->B1_UPRC,PesqPict("SB1","B1_UPRC")))
					aAdd(aForm,cObs)
					aAdd(aForm,cITMed)
					aAdd(aForm,alltrim(cCTH))
					aAdd(aForm,alltrim(cCTD))
					aAdd(aForm,alltrim(cCT1))

					//NP3-02/08/2022
					DbSelectArea("SC1")
					SC1->(DbSetOrder(2))//C1_FILIAL+C1_PRODUTO+C1_NUM+C1_ITEM
					IF SC1->(DbSeek(xFilial("SC1") + Padr(SC7->C7_PRODUTO,Len(SC1->C1_PRODUTO)) + Padr(SC7->C7_NUMSC,Len(SC1->C1_NUM)) + Padr(SC7->C7_ITEMSC,Len(SC1->C1_ITEM))   ))
						//Adicionado o ID da SC e o Solicitante -- 15/08/2022
						aAdd(aForm,SC1->C1_SOLICIT)
						aAdd(aForm,SC1->C1_NUM)
					Else
						//Adicionado o ID da SC e o Solicitante -- 15/08/2022
						aAdd(aForm,"")
						aAdd(aForm,"")
					Endif

					MontForm(2,aForm,@oForms)				

					nCount++
					SC7->(DbSkip())
				EndDo
			elseif  cTipo == 'MD'
				// nCount := 1
				DbSelectArea("SCR")
				SCR->(DbSetOrder(1))
				IF SCR->(MsSeek(xFilial("SCR") +PADR(cTipo,LEN(SCR->CR_TIPO)) + Padr(cDocumento,Len(SCR->CR_NUM))))
					
					If Empty(SCR->CR_XIDFLU)
						// IF nCount == 1
							cCdFilial  := FwCodFil()
							cCdEmpresa := FwCodEmp()
							cNmFilial  := FwFilialName()
							cNmEmpresa := FwEmpName(cCdEmpresa)

							DbSelectArea("CND")
							CND->(DbSetOrder(4))
							CND->(MsSeek(xFilial("CND") + SubStr(cDocumento,1,6)))

							//TODO: Avaliar essa modificação na linha comentada abaixo, pois não foi feito por HV. Entender o motivo pelo qual foi modificado.
							DbSelectArea("CNE")
							CNE->(DbSetOrder(6))
							CNE->(MsSeek(xFilial("CNE") + PADR(SubStr(cDocumento,1,6),len(CNE->CNE_NUMMED)) + PADR(cPlan,LEN(CNE->CNE_NUMERO)) +;
							 PADR(QTMP->DBM_ITEM,LEN(CNE->CNE_ITEM))))    
							// PADR(RIGHT(cDocumento,3),LEN(CNE->CNE_ITEM))))                                				

							DbSelectArea("CXN")
							CXN->(DbSetOrder(1))
							CXN->(MsSeek(xFilial("CXN") + PADR(CND->CND_CONTRA,LEN(CXN->CXN_CONTRA)) + PADR(CND->CND_REVISA,len(CXN->CXN_REVISA)) +;
								PADR(SubStr(cDocumento,1,6),len(CXN->CXN_NUMMED)) +	PADR(cPlan,LEN(CXN->CXN_NUMPLA)) ))
							
							DbSelectArea("CN9")
							CN9->(DbSetOrder(1))
							CN9->(MsSeek(xFilial("CN9") + CND->CND_CONTRA + CND->CND_REVISA))
							cObjt := MsMM(CN9->CN9_CODOBJ)

							cFrete := ""

							cFornec 	:= ALLTRIM(CXN->CXN_FORNEC)
							cLoja 		:= ALLTRIM(CXN->CXN_LJFORN)
							dEmissao 	:= STOD("") //CNE=>CNE_DTLANC verificar se   preenchido e se nao colocar inicializador padrao
							cContra 	:= CNE->CNE_CONTRA//ALLTRIM(CNE->CNE_CONTRA)
							cMed		:= CNE->CNE_NUMMED//ALLTRIM(CNE->CNE_NUMMED)
							cRev		:= CNE->CNE_REVISA//ALLTRIM(CNE->CNE_REVISA)
							cCCusto		:= ALLTRIM(CNE->CNE_CC)
							cProd		:= ALLTRIM(CNE->CNE_PRODUT)
							cQuant		:= Transform(CNE->CNE_QUANT,PesqPict("CNE","CNE_QUANT"))
							cPrc 		:= Transform(CNE->CNE_VLUNIT,PesqPict("CNE","CNE_VLUNIT"))
							cTot 		:= Transform(CNE->CNE_VLTOT,PesqPict("CNE","CNE_VLTOT"))
							cObs		:= CNE->CNE_XOBSM//"" //CNE->CNE_CODOBS
							cGrpAprov	:= aAprovPc[i][1]
							cItGrp		:= aAprovPc[i][10]
							cITMed		:= ALLTRIM(CNE->CNE_ITEM)
							cClvl		:= ALLTRIM(CNE->CNE_CLVL)
							cItemC		:= ALLTRIM(CNE->CNE_ITEMCT)
							cConta		:= ALLTRIM( CNE->CNE_CONTA)
							cComprador  := ""
    						cUserComp   := ""
							cEmailComp  := ""

							DbSelectArea("SA2")
							SA2->(DbSetOrder(1))
							SA2->(MsSeek(xFilial("SA2") + Padr(cFornec,Len(SA2->A2_COD)) + Padr(cLoja,Len(SA2->A2_LOJA))))

							aForm := {}
							aAdd(aForm,cCdEmpresa)
							aAdd(aForm,cNmEmpresa)
							aAdd(aForm,cCdFilial)
							aAdd(aForm,cNmFilial)
							aAdd(aForm,AllTrim(cNmUlAprv))
							aAdd(aForm,cDtUlAprv)
							aAdd(aForm,aAprovPc[i][02])
							aAdd(aForm,cDocumento)
							aAdd(aForm,DToS(dEmissao))
							aAdd(aForm,cFrete)
							aAdd(aForm,"0")
							aAdd(aForm,"")
							aAdd(aForm,SA2->A2_COD + SA2->A2_LOJA + " - " + SA2->A2_NOME)
							aAdd(aForm,SA2->A2_END)
							aAdd(aForm,SA2->A2_TEL)
							aAdd(aForm,aAprovPc[i][9])
							aAdd(aForm,aAprovPc[i][5])
							aAdd(aForm,aAprovPc[i][04])
							aAdd(aForm,cContra)
							aAdd(aForm,cMed)
							aAdd(aForm,cRev)
							aAdd(aForm,cObjt)
							aAdd(aForm,cGrpAprov)
							aAdd(aForm,cItGrp)
							aAdd(aForm,cComprador)
							aAdd(aForm,cUserComp)
							aAdd(aForm,cEmailComp)
							
							MontForm(1,aForm,@oForms)
							
						// Endif

						//Retorna descri  o do centro de custo
						cCCDesc  := Iif(Empty(cCCusto),'', Posicione("CTT",1,xFilial("CTT") + cCCusto,"CTT_DESC01"))
						cCC      := cCCusto + "-" + cCCDesc
						//Retorna descri  o da conta cont bil
						cCT1Desc := Iif(Empty(cConta),'', Posicione("CT1",1,xFilial("CT1") + cConta,"CT1_DESC01"))
						cCT1      := cConta + "-" + cCT1Desc
						//Retorna descri  o do item conta
						cCTDDesc := Iif(Empty(cItemC),'', Posicione("CTD",1,xFilial("CTD") + cItemC,"CTD_DESC01"))
						cCTD   := cItemC + "-" + cCTDDesc
						//Retorna descri  o da classe de valor
						cCTHDesc := Iif(Empty(cClvl),'', Posicione("CTH",1,xFilial("CTH") + cClvl,"CTH_DESC01"))
						cCTH    	 := cClvl + "-" + cCTHDesc

						DbSelectArea("SB1")
						SB1->(DbSetOrder(1))
						SB1->(MsSeek(xFilial("SB1") + Padr(cProd,Len(SB1->B1_COD)) ))

						cDesc := SB1->B1_DESC

						DbSelectArea("SB2")
						SB2->(DbSetOrder(1))
						SB2->(MsSeek(xFilial("SB2") +  Padr(cProd,Len(SB1->B1_COD))))
						
						aForm := {}

						aAdd(aForm,cITMed)
						aAdd(aForm,cProd)
						aAdd(aForm,cDesc)
						aAdd(aForm,cQuant)
						aAdd(aForm,cPrc)
						aAdd(aForm,cTot)
						aAdd(aForm,alltrim(cCC))
						aAdd(aForm,DtoC(SB1->B1_UCOM))
						aAdd(aForm,Transform(SB1->B1_UPRC,PesqPict("SB1","B1_UPRC")))
						aAdd(aForm,cObs)
						aAdd(aForm,cITMed)
						aAdd(aForm,alltrim(cCTH))
						aAdd(aForm,alltrim(cCTD))
						aAdd(aForm,alltrim(cCT1))

						//Adicionado o ID da SC e o Solicitante -- 15/08/2022
						aAdd(aForm,"")
						aAdd(aForm,"")
						
						MontForm(2,aForm,@oForms)				

						// nCount++
					EndIf
					
				EndIf
			EndIf
		Endif
	Next i

	IF nCount <= 1
		conOut(" HV -----> Total de registros encontrados: " + cValToChar(nCount))
		oForms := Nil
	Endif

Return

Static Function MontForm(nTp,aForm,oForms)

	If nTp == 1
		oForms["empresa"] 		:=	aForm[1]
		oForms["nomeempresa"] 	:=	aForm[2]
		oForms["filial"] 		:=	aForm[3]
		oForms["nomefilial"] 	:=	aForm[4]
		oForms["ultaprova"] 	:=	aForm[5]
		oForms["dtaprova"] 		:=	aForm[6]
		oForms["aprovador"] 	:=	aForm[7]
		oForms["pedido"] 		:=	aForm[8]
		oForms["emissao"] 		:=	aForm[9]
		oForms["tipofrete"] 	:=	aForm[10]
		oForms["totalpedido"] 	:=	aForm[11]
		oForms["codpgto"] 		:=	aForm[12]
		oForms["fornecedor"] 	:=	aForm[13]
		oForms["endereco"]		:=	aForm[14]
		oForms["telefone"]		:=	aForm[15]
		oForms["tipo"]			:=	aForm[16]
		oForms["mailuser"]		:=	aForm[17]
		oForms["nomeaprovador"] :=	aForm[18]
		oForms["contrato"] 		:=	aForm[19]
		oForms["medicao"] 		:=	aForm[20]
		oForms["revisao"] 		:=	aForm[21]
		oForms["objeto"] 		:=	aForm[22]
		oForms["grupo"]			:=	aForm[23]
		oForms["itemgrupo"]		:=	aForm[24]
		oForms["comprador"]     :=	aForm[25]
		oForms["usercomprador"] :=	aForm[26]
		oForms["emailcomprador"]:=	aForm[27]
		// DbSelectArea("CN9")
		// CN9->(DbSetOrder(1))
		// CN9->(MsSeek(aForm[3] + aForm[19] + aForm[21]))
		// DbSelectArea("RD0")
    	// RD0->(DbSetOrder(1))
    	// RD0->(MsSeek(xFilial("RD0") + CN9->CN9_XREVIS))
		oForms["revisor"]       :=	""//RD0->RD0_NOME

	else
		oForms["item___"+cValTochar(nCount)]           := aForm[1] 
		oForms["produto___"+cValTochar(nCount)]        := aForm[2] 
		oForms["descricao___"+cValTochar(nCount)]      := aForm[3] 
		oForms["quantidade___"+cValTochar(nCount)]     := aForm[4]
		oForms["valorunit___"+cValTochar(nCount)]      := aForm[5]
		oForms["valortotal___"+cValTochar(nCount)]     := aForm[6] 
		oForms["ccusto___"+cValTochar(nCount)]         := aForm[7] 
		oForms["ultcompra___"+cValTochar(nCount)]      := aForm[8] 
		oForms["vlultimacompra___"+cValTochar(nCount)] := aForm[9] 
		oForms["observacao___"+cValTochar(nCount)]     := aForm[10] 
		oForms["itemmedicao___"+cValTochar(nCount)]    := aForm[11]

		oForms["classevalor___"+cValTochar(nCount)]    := aForm[12]
		oForms["itemcontabil___"+cValTochar(nCount)]   := aForm[13]
		oForms["contacontabil___"+cValTochar(nCount)]  := aForm[14]

		//Adicionado o ID da SC e o Solicitante -- 15/08/2022
		oForms["solicitante___"+cValTochar(nCount)]    := aForm[15]
		oForms["idsolicitacao___"+cValTochar(nCount)]  := aForm[16]
    

	EndIf

RETURN


User Function CancPed(cDocumento,cCnpj,cCodfluig,lAprov,cTipo)

	Local oCanc
	Local nX
	Local aConstraints := {}
	Local oConstraints
	Local cRet := ""

	Default cCodfluig := ""
	Default lAprov := .F.


	oCanc := JsonObject():New()

	If lAprov
		oCanc['processInstanceId' ] := cCodfluig
		oCanc['cancelText' ] := "Canceled"

	Else 
		oCanc['name' ] := "ds_cancela_documento"
		oCanc['fields' ] := {"documento","empresa","filial","tipo"}
		oCanc['order'] := {"documento"}
		aConstraints := {}
		aAdd(aConstraints,JsonObject():new())
		aAdd(aConstraints,JsonObject():new())
		aAdd(aConstraints,JsonObject():new())
		aAdd(aConstraints,JsonObject():new())
		aAdd(aConstraints,JsonObject():new())
		IF !Empty(cCodFluig)
			aAdd(aConstraints,JsonObject():new())
		Endif

		For nX := 1 to Len(aConstraints)
			If nX == 1
				aConstraints[nX]['_field']        := "documento"
				aConstraints[nX]['_initialValue'] := cDocumento
				aConstraints[nX]['_finalValue']   := cDocumento
				aConstraints[nX]['_type']         := 1
				aConstraints[nX]['_likeSearch']   := .F.
			Elseif nX == 2
				aConstraints[nX]['_field']        := "empresa"
				aConstraints[nX]['_initialValue'] := FwCodEmp()
				aConstraints[nX]['_finalValue']   := FwCodEmp()
				aConstraints[nX]['_type']         := 1
				aConstraints[nX]['_likeSearch']   := .F.
			Elseif nX == 3
				aConstraints[nX]['_field']        := "filial"
				aConstraints[nX]['_initialValue'] := FwCodFil()
				aConstraints[nX]['_finalValue']   := FwCodFil()
				aConstraints[nX]['_type']         := 1
				aConstraints[nX]['_likeSearch']   := .F.
			Elseif nX == 4
				aConstraints[nX]['_field']        := "tipo"
				aConstraints[nX]['_initialValue'] := cTipo
				aConstraints[nX]['_finalValue']   := cTipo
				aConstraints[nX]['_type']         := 1
				aConstraints[nX]['_likeSearch']   := .F.
			Elseif nX == 5
				aConstraints[nX]['_field']        := "usuario"
				aConstraints[nX]['_initialValue'] := FwGetUserName(RetCodUsr())
				aConstraints[nX]['_finalValue']   := FwGetUserName(RetCodUsr())
				aConstraints[nX]['_type']         := 1
				aConstraints[nX]['_likeSearch']   := .F.
			Else
				aConstraints[nX]['_field']        := "codfluig"
				aConstraints[nX]['_initialValue'] := cCodfluig
				aConstraints[nX]['_finalValue']   := cCodfluig
				aConstraints[nX]['_type']         := 1
				aConstraints[nX]['_likeSearch']   := .F.
			EndIf
		Next nX
		oConstraints := aConstraints
		oCanc['constraints'] := oConstraints
	EndIf

	If lAprov
		cRet := U_ConFluig(oCanc:ToJson(),"/api/public/2.0/workflows/cancelInstance")
		Sleep(500)
	Else
		cRet := U_ConFluig(oCanc:ToJson(),"/api/public/ecm/dataset/datasets")
		Sleep(500)
	EndIf
Return cRet


#Include "Protheus.ch"

/*/{Protheus.doc} ConFluig
Envio de payload JSON para Webhook do Hubot/n8n.
@type    User Function
@author  Suporte
@since   03/08/2026
@param   cBody,  Character, JSON que será enviado no corpo da requisição.
@param   cPath,  Character, Caminho complementar (opcional/reservado).
@param   nPath,  Numeric,   1 = Aprov. Compra | Outros = Aprov. Pedido.
@return  cRet,   Character, Resposta (String/JSON) retornada pelo Webhook.
/*/
User Function ConFluig(cBody, cPath, nPath)
    Local cBaseUrl  := AllTrim(SuperGetMV("HV_WHCOMPRA", .F., "https://workflows.hubot.app.br/webhook")) 
    Local cEndpoint := ""
    Local cRet      := ""
    Local aHeader   := {}
    Local oRest     := Nil

    Default cBody   := ""
    Default cPath   := ""
    Default nPath   := 1

    // Trata a barra no final do parâmetro do MV
    If Right(cBaseUrl, 1) == "/"
        cBaseUrl := SubStr(cBaseUrl, 1, Len(cBaseUrl) - 1)
    EndIf

    // Define o endpoint de acordo com o parâmetro nPath
    If nPath == 1
        cEndpoint := "/aprovacaodepedido"
    Else
        cEndpoint := "/aprovacaodecompra"
    EndIf

    // Instancia a classe com a URL base e define o caminho no SetPath
    oRest := FWRest():New(cBaseUrl)
    oRest:SetPath(cEndpoint)
	oRest:SetPostParams(cBody)

    // Configura os cabeçalhos da requisição
    AAdd(aHeader, "Content-Type: application/json; charset=utf-8")

    // ATENÇÃO: Passar o cBody como 2º parâmetro no Post resolve o 'type mismatch on +'
    If oRest:Post(aHeader)
        cRet := oRest:GetResult()
    Else
        cRet := oRest:GetLastError()
        If Empty(cRet)
            cRet := oRest:GetResult()
        EndIf
    EndIf

Return cRet
// User Function ConFluig(cBody,cPath)
// 	Local cRet            := ""
// 	local oBody           := NIL
// 	Local cAccessToken    := SuperGetMV("MV_FTOKENA",.F.,"") //
// 	Local cTokenSecret    := SuperGetMV("MV_FTOKENS",.F.,"")  //
// 	Local cURL            := SuperGetMV("MV_JFLGURL",.T.,"") //
// 	Local cConsumerKey    := SuperGetMV("MV_FCKEY",.F.,"")
// 	Local cConsumerSecret := SuperGetMV("MV_FCSECRE",.F.,"")
// 	local cx_url	      := cURL + cPath

// 	// u_logInt("1", cPath, cBody)

// 	cAccess    := cURL+'/portal/api/rest/oauth/access_token'
// 	cRequest   := cURL+'/portal/api/rest/oauth/request_token'
// 	cAuthorize := cURL+'/portal/api/rest/oauth/authorize'

// 	ConOut("------------HV --------- Iniciando integracao com FLUIG ")
// 	Sleep(500)
// 	oUrl    := FWoAuthURL():New( cRequest , cAuthorize , cAccess )
// 	//Sleep(500)

// 	oClient := fwOAuthClient():new(cConsumerKey, cConsumerSecret, oUrl, cx_url)
// 	//Sleep(500)

// 	oClient:cOAuthVersion   := "1.0"
// 	oClient:SetContentType("application/json")
// 	oClient:setMethodSignature("HMAC-SHA1")
// 	oClient:setToken(cAccessToken)
// 	oClient:setSecretToken(cTokenSecret)
// 	oClient:makeSignBaseString("POST", cx_url)
// 	oClient:MakeSignature()

// 	fwJsonDeserialize(cBody, @oBody)

// 	// cRet := oClient:Post(cx_url, "", cBody )

// 	IF SuperGetMv("HV_LOGAPI",,.T.)
// 		ConOut("------------HV --------- Retorno Fluig: " + iif(cRet == NIL, "Nil",cRet))
// 	Endif

// 	ConOut("------------HV --------- Finalizando integracao com FLUIG ")

// Return cRet

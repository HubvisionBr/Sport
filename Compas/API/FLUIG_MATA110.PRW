#include 'RESTFUL.CH'
#include 'TOTVS.ch'
#include 'TOPCONN.CH'
#include 'TBICONN.CH'
#Include "FWMVCDEF.CH"
#include 'parmtype.ch'


WSRESTFUL FLUIG_MATA110 DESCRIPTION "API INTEGRACAO FLUIG COM SOLICITACAO DE COMPRAS totvs"
	WSMETHOD POST DESCRIPTION "Verifica Dados do POST" WSSYNTAX "/fluig_mata110"
END WSRESTFUL

WSMETHOD POST WSSERVICE FLUIG_MATA110
	Local lPost := .T.
	Local aCabec := {}
	Local aItens := {}
	Local aLinha := {}
	Local nX     := 0
	Local Json := Nil
	Local cBody := ::GetContent()
	Local nQtde := 0
	Local nPreco := 0
	Local oResponse
	// local aArea    	    := GetArea()
	// local aAreaSC1 	    := SC1->(GetArea())
	local cDoc          := GetSXENum("SC1","C1_NUM")
	Private aAutoErro      := {}
	Private lMsErroAuto    := .F.
	Private lMsHelpAuto    := .T.
	Private lAutoErrNoFile := .T.
	Private cMensag        := ""

	// u_logInt("2", "FLUIG_MATA110", Self:GetContent())

	::SetContentType("application/json")

	FWJsonDeserialize(cBody, @Json)

	aadd(aCabec,{"C1_FILIAL",  xFilial("SC1")}) //Estou enviando o TenaltId. Precisa dessa chave da filial?
	aadd(aCabec,{"C1_SOLICIT", FwNoAccent(DecodeUTF8(Json:solicitante, "cp1252")) })
	aadd(aCabec,{"C1_EMISSAO", CTOD(Json:dtsolicitacao)})

	SC1->(dbSetOrder(1))
	while SC1->(dbSeek(xFilial("SC1")+cDoc))
		ConfirmSX8()
		cDoc := GetSXENum("SC1","C1_NUM")
	endDo

	aadd(aCabec,{"C1_NUM" ,cDoc})

	For nX := 1 To LEN(Json:itens)
		aLinha  := {}

		If ValType(Json:itens[nX]:quantidade) == "C"
			nQtde := val(Json:itens[nX]:quantidade)
		Else
			nQtde := Json:itens[nX]:quantidade
		EndIf

		If ValType(Json:itens[nX]:preco) == "C"
			nPreco := val(Json:itens[nX]:preco)
		Else
			nPreco := Json:itens[nX]:preco
		EndIf

		aadd(aLinha,{"C1_ITEM"   	, STRZERO(nX,4,0), Nil})
		aadd(aLinha,{"C1_PRODUTO"	, Json:itens[nX]:produto, Nil})
		aadd(aLinha,{"C1_UM" 		, Json:itens[nX]:unidade, Nil})
		aadd(aLinha,{"C1_QUANT"  	, nQtde, Nil})
		aadd(aLinha,{"C1_PRECO"  	, nPreco, Nil})
		aadd(aLinha,{"C1_DATPRF"    , CTOD(Json:itens[nX]:dtnecessidade), Nil})
		aadd(aLinha,{"C1_RATEIO"    , "2", Nil})
		aadd(aLinha,{"C1_CC"        , Json:itens[nX]:centrocusto, Nil})
		aadd(aLinha,{"C1_XORIG"     , "F", Nil})
		aadd(aLinha,{"C1_XCODFLU"   , Json:solicitacao, Nil})
		aadd(aLinha,{"C1_XITEMF"    , Json:itens[nX]:id, Nil})
		aadd(aLinha,{"C1_OBS"   	, FwNoAccent(DecodeUTF8(Json:itens[nX]:observacao, "cp1252")), Nil})

		//Novos campos add
		// IF !EMPTY(Json:itens[nX]:contacontabil)
		// 	aadd(aLinha,{"C1_CONTA"       , Json:itens[nX]:contacontabil, Nil})
		// ENDIF
		// IF !EMPTY(Json:itens[nX]:itemcontabil)
		// 	aadd(aLinha,{"C1_ITEMCTA"       , Json:itens[nX]:itemcontabil, Nil})
		// ENDIF
		// IF !EMPTY(Json:itens[nX]:classevalor)
		// 	aadd(aLinha,{"C1_CLVL"       , Json:itens[nX]:classevalor, Nil})
		// ENDIF

		IF !EMPTY(Json:itens[nX]:grupocomprador)
			aadd(aLinha,{"C1_XGRUPCO"   , Json:itens[nX]:grupocomprador, Nil})
			/*
			*TODOS
			* Não está gravando no campo padrão porque será?
			*/
			aadd(aLinha,{"C1_GRUPCOM"   , Json:itens[nX]:grupocomprador, Nil})
		ENDIF
		

		aadd(aItens,aLinha)

	Next


	MSExecAuto({|x,y,z,a,b,c| Mata110(x,y,z,a,b,c)},aCabec,aItens, 3,Nil,Nil,Nil,Nil )

	If lMsErroAuto
		aAutoErro := GETAUTOGRLOG()
		// cMensag   := AllTrim(U_SODE991(aAutoErro))

		SetRestFault(400, "Erro ao criar solicitacao de compra, Error:"+cMensag)
		lPost := .F.

	Else

		oResponse := JsonObject():new()
		oResponse['status'] := "success"
		oResponse['code'] 	:= cDoc
		oResponse['message'] 	:= "Solicitacao de Compra incluida com sucesso. N: "  + xFilial("SC1") + cDoc

		::SetResponse(EncodeUTF8(oResponse:toJson()))


	EndIf


Return lPost

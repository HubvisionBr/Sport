#Include 'Protheus.ch'
#Include 'FWMVCDEF.ch'
#Include 'RestFul.CH'
#INCLUDE "TOTVS.CH"
#INCLUDE "TopConn.ch"
#INCLUDE 'COLORS.CH'
#INCLUDE 'FONT.CH'
#INCLUDE 'RWMAKE.CH'
#INCLUDE "TBICONN.CH"
#include "parmtype.ch"
#INCLUDE "FWADAPTEREAI.CH"
#Include "TBICODE.CH"
#include 'fileio.ch'

#DEFINE ENTER chr(10)+chr(13)


//Inicio da declaração da estrutura do Webservice;
WSRESTFUL HVREST DESCRIPTION "Serviço REST para requisições ao ERP Protheus"
	
	WSDATA cWsFuncao AS CHARACTER OPTIONAL

	WSMETHOD POST   LISTA     	DESCRIPTION "Retorna lista de registros da empresa - generico. Será preciso informar no BODY uma lista de registros com a informação da empresa, filial, select, campos de retorno e delimitador."  WSSYNTAX "HVREST/LISTA" PATH "/LISTA"

END WSRESTFUL

//Methodo post para execução de querys genericas
WSMETHOD POST LISTA WSSERVICE HVREST

	Local cJson     := Self:GetContent()
	Local oJson		:= JsonObject():New()
	Local oRet      := JsonObject():New()
	Local lRet		:= .t.
	Local oJsonRet := nil
	Local oItem := nil
	Local aRetorno  := {}
	Local cObjJson := ''
	Local cQuery 	:= ''
	Local cAlias    := GetNextAlias()
	Local aWsFiltro := {} //strtokarr(::cWsFiltro, ::cWsDelimitador)
	Local nX0 := 0

	PRIVATE cError := ""
	//PRIVATE oLastError := ErrorBlock({|e| cError := "ERROR: " + e:Description +chr(10)+ e:ErrorStack})
	PRIVATE oLastError := ErrorBlock( { |e| cError := e:ErrorStack, Break(e) } )

	::SetContentType("application/json")

    If Type('cFilAnt') == 'U'
		cSetEmp := SubStr(Self:GetHeader('tenantId'),1,2)
		cSetFil := SubStr(Self:GetHeader('tenantId'),4)
		RpcSetEnv(cSetEmp,cSetFil)
	EndIF


	BEGIN SEQUENCE

		oJson := JsonObject():New()

		oRet := oJson:FromJson(cJson)
		cObjJson := oJson:GetNames()[1]
		if ValType(oJson[cObjJson]) == "C"

			if empty(oJson[cObjJson]) //.or. empty(aCampos[2]) //.or. empty(aCampos[3])

				cError := 'É preciso informar a query para devida execução da rotina! Ex.: {"query": "select * from CTT010 where ... "}"'
			
			endif

			cQuery := ChangeQuery(oJson[cObjJson]) //query para execução

			TcQuery cQuery Alias (cAlias) New
			(cAlias)->(DbGoTop())

			//listando todos os campos da query
			aWsFiltro := (cAlias)->(DBSTRUCT())


			(cAlias)->(DbGoTop())
			nX0 := 0
			Do While !(cAlias)->(EOF())

				oItem := JsonObject():new()

				for nX0 := 1 to len(aWsFiltro)

					if !empty(aWsFiltro[nx0][1])

						IF Valtype((cAlias)->&(aWsFiltro[nx0][1])) == "C"

							oItem[aWsFiltro[nx0][1]] := EncodeUTF8((cAlias)->&(aWsFiltro[nx0][1]))

						else

							oItem[aWsFiltro[nx0][1]] := (cAlias)->&(aWsFiltro[nx0][1])

						Endif

					endif

				next nX0

				AADD(aRetorno,oItem)

				(cAlias)->(dbskip())

			enddo
			(cAlias)->(DbCloseArea())

			oJsonRet := JsonObject():new()
			oJsonRet:Set(aRetorno)

			::SetResponse('{"LISTA":' +oJsonRet:ToJson()+ '}')

		Else

			cError := "Falha ao executar JsonObject. Erro: objeto array nao identificado no bory do post"

		endif
		FreeObj(oJson)
		ErrorBlock(oLastError)

	END SEQUENCE

	if !empty(cError)

		self:setStatus(400)
		SetRestFault(400, EncodeUTF8( cError ), .T.,400)
		lRet := .f.
	endif

Return lRet

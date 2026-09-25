//Bibliotecas
#Include "Protheus.ch"

//Constantes
#Define STR_PULA	Chr(13)+Chr(10)

//Variáveis estáticas
Static cDirTmp := GetTempPath()

/*/{Protheus.doc} HVSPT01
Função Genérica para execauto
@type function
@author Atilio
@since 25/10/2015
@version 1.0
	@example
	u_HVSPT01()
/*/
User Function HVSPT01()
	Local aArea := GetArea()
	//Dimensões da janela
	Local nJanAltu := 180
	Local nJanLarg := 650
	//Objetos da tela
	Local oGrpPar
	Local oGrpAco
	Local oBtnSair
	Local oBtnImp
	Local oBtnObri
	Local oBtnArq
	Local cAviso := ""
	Private aIteTip := {}
	Private oSayArq, oGetArq, cGetArq := Space(99)
	Private oSayTip, oCmbTip, cCmbTip := ""
	Private oSayCar, oGetCar, cGetCar := ';'
	Private oDlgPvt
	
	//Inserindo as opções disponí­veis no Carga Dados Genérico
	aIteTip := {;
		"01=Baixa a Pagar",;
		"02=Baixa a Receber";
	}
	cCmbTip := aIteTip[1]
	
	//Mostrando um aviso sobre a importação
	cAviso := "HBSPT01: Rotinas Automáticas - Genérico v1.0"+STR_PULA
	cAviso += "--"+STR_PULA
	cAviso += "Para campos Numéricos com separação de decimal, utilize o caracter '.'. Por exemplo: 5.20;"+STR_PULA
	cAviso += "Para campos do tipo Data, utilize ou o padrão YYYYMMDD ou o DD/MM/YYYY. Por exemplo: 20151025 ou 25/10/2015;"+STR_PULA
	cAviso += "--"+STR_PULA
	cAviso += "A rotina está preparada para importar os seguintes cadastros:"+STR_PULA
	cAviso += " Seq- Rotina  "+STR_PULA
	cAviso += " 01 - FINA080 "+STR_PULA
	cAviso += " 02 - FINA070 "+STR_PULA
	cAviso += "--"+STR_PULA
	cAviso += " O caracter ';' (ponto e ví­rgula), nunca pode estar no fim da linha!"+STR_PULA
	Aviso('Atenção', cAviso, {'Ok'}, 03)
	
	//Criando a janela
	DEFINE MSDIALOG oDlgPvt TITLE "Carga Dados - Genérico" FROM 000, 000  TO nJanAltu, nJanLarg COLORS 0, 16777215 PIXEL
		//Grupo Parâmetros
		@ 003, 003 	GROUP oGrpPar TO 060, (nJanLarg/2) 	PROMPT "Parâmetros: " 		OF oDlgPvt COLOR 0, 16777215 PIXEL
			//Caminho do arquivo
			@ 013, 006 SAY        oSayArq PROMPT "Arquivo:"                  SIZE 060, 007 OF oDlgPvt PIXEL
			@ 010, 070 MSGET      oGetArq VAR    cGetArq                     SIZE 240, 010 OF oDlgPvt PIXEL
			oGetArq:bHelp := {||	ShowHelpCpo(	"cGetArq",;
									{"Arquivo CSV que será importado."+STR_PULA+"Exemplo: C:\teste.CSV"},2,;
									{},2)}
			@ 010, 311 BUTTON oBtnArq PROMPT "..."      SIZE 008, 011 OF oDlgPvt ACTION (fPegaArq()) PIXEL
			
			//Tipo de Importação
			@ 028, 006 SAY        oSayTip PROMPT "Tipo Importação:"          SIZE 060, 007 OF oDlgPvt PIXEL
			@ 025, 070 MSCOMBOBOX oCmbTip VAR    cCmbTip ITEMS aIteTip       SIZE 100, 010 OF oDlgPvt PIXEL
			oCmbTip:bHelp := {||	ShowHelpCpo(	"cCmpTip",;
									{"Tipo de Importação que será processada."+STR_PULA+"Exemplo: 1 = Bancos"},2,;
									{},2)}
			
			//Caracter de Separação do CSV
			@ 043, 006 SAY        oSayCar PROMPT "Carac.Sep.:"               SIZE 060, 007 OF oDlgPvt PIXEL
			@ 040, 070 MSGET      oGetCar VAR    cGetCar                     SIZE 030, 010 OF oDlgPvt PIXEL VALID fVldCarac()
			oGetArq:bHelp := {||	ShowHelpCpo(	"cGetCar",;
									{"Caracter de separação no arquivo."+STR_PULA+"Exemplo: ';'"},2,;
									{},2)}
			
		//Grupo Ações
		@ 063, 003 	GROUP oGrpAco TO (nJanAltu/2)-3, (nJanLarg/2) 	PROMPT "Ações: " 		OF oDlgPvt COLOR 0, 16777215 PIXEL
		
			//Botões
			@ 070, (nJanLarg/2)-(63*1)  BUTTON oBtnSair PROMPT "Sair"          SIZE 60, 014 OF oDlgPvt ACTION (oDlgPvt:End()) PIXEL
			@ 070, (nJanLarg/2)-(63*2)  BUTTON oBtnImp  PROMPT "Importar"      SIZE 60, 014 OF oDlgPvt ACTION (Processa({|| fConfirm(1) }, "Aguarde...")) PIXEL
			@ 070, (nJanLarg/2)-(63*3)  BUTTON oBtnObri PROMPT "Camp.Obrig."   SIZE 60, 014 OF oDlgPvt ACTION (Processa({|| fConfirm(2) }, "Aguarde...")) PIXEL
	ACTIVATE MSDIALOG oDlgPvt CENTERED
	
	RestArea(aArea)
Return
/*---------------------------------------------------------------------------------*
 | Func.: fVldCarac                                                                |
 | Autor: Daniel Atilio                                                            |
 | Data:  25/10/2015                                                               |
 | Desc:  Função que valida o caracter de separação digitado                       |
 *---------------------------------------------------------------------------------*/
Static Function fVldCarac()
	Local lRet := .T.
	Local cInvalid := "'./\"+'"'
	
	//Se o caracter estiver contido nos que não podem, retorna erro
	If cGetCar $ cInvalid
		lRet := .F.
		MsgAlert("Caracter inválido, ele não estar contido em <b>"+cInvalid+"</b>!", "Atenção")
	EndIf
Return lRet
/*---------------------------------------------------------------------------------*
 | Func.: fPegaArq                                                                 |
 | Autor: Daniel Atilio                                                            |
 | Data:  25/10/2015                                                               |
 | Desc:  Função responsável por pegar o arquivo de importação                     |
 *---------------------------------------------------------------------------------*/
Static Function fPegaArq()
	Local cArqAux := ""
	cArqAux := cGetFile( "Arquivo Texto | *.csv",;				//Máscara
							"Arquivo...",;						//Tí­tulo
							,;										//Número da máscara
							,;										//Diretório Inicial
							.F.,;									//.F. == Abrir; .T. == Salvar
							GETF_LOCALHARD,;						//Diretório full. Ex.: 'C:\TOTVS\arquivo.xlsx'
							.F.)									//Não exibe diretório do servidor
								
	//Caso o arquivo não exista ou estiver em branco ou não for a extensão txt
	If Empty(cArqAux) .Or. !File(cArqAux) .Or. (SubStr(cArqAux, RAt('.', cArqAux)+1, 3) != "txt" .And. SubStr(cArqAux, RAt('.', cArqAux)+1, 3) != "csv")
		MsgStop("Arquivo <b>inválido</b>!", "Atenção")
		
	//Senão, define o get
	Else
		cGetArq := PadR(cArqAux, 99)
		oGetArq:Refresh()
	EndIf
Return
/*---------------------------------------------------------------------------------*
 | Func.: fConfirm                                                                 |
 | Autor: Daniel Atilio                                                            |
 | Data:  25/10/2015                                                               |
 | Desc:  Função de confirmação da tela principal                                  |
 *---------------------------------------------------------------------------------*/
Static Function fConfirm(nTipo)
	Local nModBkp			:= nModulo
	Local aAux				:= {}
	Local nAux				:= 0
	Local cAux				:= ""
	Local cFunBkp			:= FunName()
	Default nTipo			:= 2
	Private cRotina		:= ""
	Private cTabela		:= ""
	Private cCampoChv		:= ""
	Private cFilialTab	:= ""
	Private nTotalReg		:= 0
	Private cAliasTmp		:= "TMP_"+RetCodUsr()
	Private oBrowChk
	Private cFiles
	Private cMark			:= "OK"
	Private aCampos		:= {}
	Private aStruTmp		:= {}
	Private aHeadImp		:= {}
	Private cCampTipo		:= ""
	Private lChvProt		:= .F.
	Private lFilProt		:= .F.
	Private cLinhaCab		:= ""
	
	//Bancos
	If cCmbTip == "01"
		cRotina	:= "MSExecAuto({|x, y| FINA080(x, y)}, aDados, 3) "
		cTabela	:= "SE2"
		cCampoChv	:= ""
		cFilialTab	:= FWxFilial(cTabela)
		nModulo	:= 6
		SetFunName("FINA080")
	
	//Clientes
	ElseIf cCmbTip == "02"
		cRotina	:= "MSExecAuto({|x, y| FINA070(x, y)}, aDados, 3) "
		cTabela	:= "SE1"
		cCampoChv	:= ""
		cFilialTab	:= FWxFilial(cTabela)
		nModulo	:= 5
		SetFunName("FINA070")


	//Opção inválida
	Else
		nModulo	:= nModBkp
		MsgStop("Opção <b>Inválida</b>!", "Atenção")
		Return
	EndIf
	
	//Importação dos dados
	If nTipo == 1
		//Se o arquivo existir
		If File(cGetArq)
			//Abrindo o arquivo
			Ft_FUse(cGetArq)
			nTotalReg := Ft_FLastRec()
			
			//Se o total de registros for menor que 2, arquivo inválido
			If nTotalReg < 2
				MsgAlert("Arquivo inválido, possui menos que <b>2</b> linhas!", "Atenção")
				
			//Senão, chama a tela de observação e depois a importação
			Else
				//Monta tabela temporária
				fMontaTmp()
				
				//Pegando o cabeçalho
				cLinhaCab := Ft_FReadLn()
				cLinhaCab := Iif(SubStr(cLinhaCab, Len(cLinhaCab)-1, 1) == ";", SubStr(cLinhaCab, 1, Len(cLinhaCab)-1), cLinhaCab)
				aAux := Separa(cLinhaCab, cGetCar)
				Ft_FSkip()
				
				//Percorrendo o aAux e adicionando no array
				For nAux := 1 To Len(aAux)
					IF Alltrim(aAux[nAux]) $ "AUTDTBAIXA,AUTDTCREDITO"
						//Adiciona na grid
						aAdd(aHeadImp, {	aAux[nAux],;								//Campo
											 "D",;			//Tipo
											.F.})
					ElseIf Alltrim(aAux[nAux]) $ "AUTJUROS,AUTVLRPG,AUTVLREC"
						//Adiciona na grid
						aAdd(aHeadImp, {	aAux[nAux],;								//Campo
											 "N",;			//Tipo
											.F.})
					Else
						cAux := GetSX3Cache(aAux[nAux], 'X3_TIPO')
					
						//Se o tí­tulo estiver em branco, quer dizer que o campo não existe, então é um campo reservado do execauto (como o LINPOS)
						If Empty(GetSX3Cache(aAux[nAux], 'X3_TITULO'))
							cCampTipo += aAux[nAux]+";"
						EndIf
						
						//Adiciona na grid
						aAdd(aHeadImp, {	aAux[nAux],;								//Campo
											Iif(Empty(cAux), ' ', cAux),;			//Tipo
											.F.})
					Endif										//Excluí­do
				Next
				
				//Chama a tela de observação para preenchimento das informações auxiliares
				If fTelaObs(!Empty(cCampoChv))
				
					//Chama a rotina de importação
					fImport()
					
					//Se houve erros na rotina
					(cAliasTmp)->(DbGoTop())
					If ! (cAliasTmp)->(EoF())
						fTelaErro()
						
					//Senão, mostra mensagem de sucesso
					Else
						MsgInfo("Importação finalizada com Sucesso!", "Atenção")
					EndIf
				EndIf
				
				//Fechando a tabela e excluindo o arquivo temporário
				(cAliasTmp)->(DbCloseArea())
				fErase(cAliasTmp + GetDBExtension())
			EndIf
			Ft_FUse()
		
		//Senão, mostra erro
		Else
			MsgAlert("Arquivo inválido / não encontrado!", "Atenção")
		EndIf
		
	//Geração de arquivo com cabeçalho dos campos obrigatórios
	ElseIf nTipo == 2
		fObrigat()
	EndIf
	
	nModulo := nModBkp
	SetFunName(cFunBkp)
Return
/*---------------------------------------------------------------------------------*
 | Func.: fObrigat                                                                 |
 | Autor: Daniel Atilio                                                            |
 | Data:  25/10/2015                                                               |
 | Desc:  Função que gera os campos obrigatórios em CSV / TXT                      |
 *---------------------------------------------------------------------------------*/
Static Function fObrigat()
	Local aAreaX3		:= SX3->(GetArea())
	Local cConteud	:= ""
	Local cCaminho	:= cDirTmp
	Local cArquivo	:= "obrigatorio."
	Local cExtensao	:= ""
    Local aBxSE1    := {}
    Local aBxSE2    := {}
    Local nAux      := 0
	
	IF cTabela == "SE1"

        aBxSE1 := {{"E1_PREFIXO"  ,"   "                ,Nil    },;
            {"E1_NUM"      ,"200      "            ,Nil    },;
            {"E1_PARCELA"  ," "                    ,Nil    },;
            {"E1_TIPO"     ,"NF "                  ,Nil    },;
            {"AUTMOTBX"    ,"NOR"                  ,Nil    },;
            {"AUTBANCO"    ,"001"                  ,Nil    },;
            {"AUTAGENCIA"  ,"00001"                ,Nil    },;
            {"AUTCONTA"    ,"0000000001"           ,Nil    },;
            {"AUTDTBAIXA"  ,dDataBase              ,Nil    },;
            {"AUTDTCREDITO",dDataBase              ,Nil    },;
            {"AUTHIST"     ,"BAIXA TESTE"          ,Nil    },;
            {"AUTJUROS"    ,0                      ,Nil,.T.},;
            {"AUTVALREC"   ,700                    ,Nil    }}

        For nAux := 1 To Len(aBxSE1)
    	    cConteud += Alltrim(aBxSE1[nAux,1])+cGetCar
		Next nAux

	ElseIf cTabela == "SE2"

        aBxSE2 := {{"E2_PREFIXO"  ,"   "                ,Nil    },;
            {"E2_NUM"      ,"200      "            ,Nil    },;
            {"E2_PARCELA"  ," "                    ,Nil    },;
            {"E2_TIPO"     ,"NF "                  ,Nil    },;
            {"AUTMOTBX"    ,"NOR"                  ,Nil    },;
            {"AUTBANCO"    ,"001"                  ,Nil    },;
            {"AUTAGENCIA"  ,"00001"                ,Nil    },;
            {"AUTCONTA"    ,"0000000001"           ,Nil    },;
            {"AUTDTBAIXA"  ,dDataBase              ,Nil    },;
            {"AUTDTCREDITO",dDataBase              ,Nil    },;
            {"AUTHIST"     ,"BAIXA TESTE"          ,Nil    },;
            {"AUTJUROS"    ,0                      ,Nil,.T.},;
            {"AUTVLRPG "   ,700                    ,Nil    }}

        For nAux := 1 To Len(aBxSE2)
    	    cConteud += Alltrim(aBxSE2[nAux,1])+cGetCar
		Next nAux
    Endif

	//Se escolher txt
	If MsgYesNo("Deseja gerar com a extensão <b>txt</b>?", "Atenção")
		cExtensao := "txt"
		
	//Senão, será csv
	Else
		cExtensao := "csv"
	EndIf
	
	//Gera o arquivo
	MemoWrite(cCaminho+cArquivo+cExtensao, cConteud)
	
	//Tentando abrir o arquivo
	nRet := ShellExecute("open", cArquivo+cExtensao, "", cCaminho, 1)
	
	//Se houver algum erro
	If nRet <= 32
		MsgStop("Não foi possí­vel abrir o arquivo <b>"+cCaminho+cArquivo+cExtensao+"</b>!", "Atenção")
	EndIf 
	
	RestArea(aAreaX3)
Return
/*---------------------------------------------------------------------------------*
 | Func.: fMontaTmp                                                                |
 | Autor: Daniel Atilio                                                            |
 | Data:  25/10/2015                                                               |
 | Desc:  Função que monta a estrutura da tabela temporária com os erros           |
 *---------------------------------------------------------------------------------*/
Static Function fMontaTmp()
	Local aArea := GetArea()
	
	//Se tiver aberto a temporária, fecha e exclui o arquivo
	If Select(cAliasTmp) > 0
		(cAliasTmp)->(DbCloseArea())
	EndIf
	fErase(cAliasTmp + GetDBExtension())
	
	//Adicionando a Estrutura (Campo, Tipo, Tamanho, Decimal)
	aStruTmp:={}
	aAdd(aStruTmp,{	"TMP_SEQ",		"C",	010,						0})
	aAdd(aStruTmp,{	"TMP_LINHA",	"N",	018,						0})
	aAdd(aStruTmp,{	"TMP_ARQ",		"C",	250,						0})
	
	//Criando tabela temporária
	cFiles := CriaTrab( aStruTmp, .T. )             
	dbUseArea( .T., "DBFCDX", cFiles, cAliasTmp, .T., .F. )
	
	//Setando os campos que serão mostrados no MsSelect
	aCampos := {}
	aAdd(aCampos,{	"TMP_SEQ",		,	"Sequencia",		"@!"})
	aAdd(aCampos,{	"TMP_LINHA",	,	"Linha Erro",		""})
	aAdd(aCampos,{	"TMP_ARQ",		,	"Arquivo Log.",	""})
	
	RestArea(aArea)
Return
/*---------------------------------------------------------------------------------*
 | Func.: fTelaObs                                                                 |
 | Autor: Daniel Atilio                                                            |
 | Data:  25/10/2015                                                               |
 | Desc:  Função de observações antes da importação                                |
 *---------------------------------------------------------------------------------*/
Static Function fTelaObs(lAtivChav)
	Local lRet := .F.
	//Dimensões da janela
	Local nJanAltu := 500
	Local nJanLarg := 700
	//Objetos da tela
	Local oGrpDad
	Local oGrpCam
	Local oGrpOpc
	Local oGrpAco
	Local oBtnConf
	Local oBtnCanc
	//Janela
	Local oDlgObs
	//Radios - Chave
	Local oSayChave, oRadChave, nRadChave := 1
	//Radios - Filial
	Local oSayFilial, oRadFilial, nRadFilial := 1
	//Campos no grupo de Dados
	Local oSayTab, oGetTab, cGetTab := cTabela
	Local oSayCam, oGetCam, cGetCam := cCampoChv
	Local oSayFil, oGetFil, cGetFil := cFilialTab
	Local oSayRot, oGetRot, cGetRot := cRotina
	//Grid
	Private oMsNew
	Private aHeadNew := {}
	Private aColsNew := aClone(aHeadImp)
	Default lAtivChav := .T.
	
	//Setando o cabeçalho
	//Cabeçalho ...	Titulo				Campo			Mask		Tamanho	Dec		Valid				Usado	Tip		F3	CBOX
	aAdd(aHeadNew,{	"Campo",			"XX_CAMP",		"@!",		10,			0,		".F.",				".F.",	"C",	"",	""})
	aAdd(aHeadNew,{	"Tipo",			"XX_TIPO",		"@!",		1,			0,		"u_HVSPTA01()",	".T.",	"C", 	"",	"C;N;L;D;M",	"C=Caracter;N=Numérico;L=Lógico;D=Data;M=Memo"})
	
	//Criando a janela
	DEFINE MSDIALOG oDlgObs TITLE "Observações" FROM 000, 000  TO nJanAltu, nJanLarg COLORS 0, 16777215 PIXEL STYLE DS_MODALFRAME
		//Grupo Dados
		@ 003, 003 	GROUP oGrpDad TO 055, (nJanLarg/2) 	PROMPT "Dados: " 		OF oDlgObs COLOR 0, 16777215 PIXEL
			//Tabela
			@ 010, 005 SAY   oSayTab PROMPT "Tabela:"  SIZE 040, 011 OF oDlgObs COLORS 0, 16777215 PIXEL
			@ 017, 005 MSGET oGetTab VAR    cGetTab    SIZE 040, 010 OF oDlgObs PIXEL
			oGetTab:lActive := .F.
			
			//Campo Chave
			@ 010, 121 SAY   oSayCam PROMPT "Campo Chave:"  SIZE 040, 011 OF oDlgObs COLORS 0, 16777215 PIXEL
			@ 017, 121 MSGET oGetCam VAR    cGetCam    SIZE 040, 010 OF oDlgObs PIXEL
			oGetCam:lActive := .F.
			
			//Filial
			@ 010, 237 SAY   oSayFil PROMPT "Filial Atual:"  SIZE 040, 011 OF oDlgObs COLORS 0, 16777215 PIXEL
			@ 017, 237 MSGET oGetFil VAR    cGetFil    SIZE 040, 010 OF oDlgObs PIXEL
			oGetFil:lActive := .F.
			
			//Rotina
			@ 031, 005 SAY   oSayRot PROMPT "Rotina:"  SIZE 040, 011 OF oDlgObs COLORS 0, 16777215 PIXEL
			@ 038, 005 MSGET oGetRot VAR    cGetRot    SIZE 272, 010 OF oDlgObs PIXEL
			oGetRot:lActive := .F.
			
		//Grupo Campos
		@ 058, 003 	GROUP oGrpCam TO 180, (nJanLarg/2) 	PROMPT "Campos: " 	OF oDlgObs COLOR 0, 16777215 PIXEL
			oMsNew := MsNewGetDados():New(	058+12,;									//nTop
    											006,;										//nLeft
    											177,;										//nBottom
    											(nJanLarg/2)-6,;							//nRight
    											GD_INSERT+GD_DELETE+GD_UPDATE,;			//nStyle
    											"AllwaysTrue()",;							//cLinhaOk
    											,;											//cTudoOk
    											"",;										//cIniCpos
    											{"XX_TIPO"},;								//aAlter
    											,;											//nFreeze
    											999,;										//nMax
    											,;											//cFieldOK
    											,;											//cSuperDel
    											,;											//cDelOk
    											oDlgObs,;									//oWnd
    											aHeadNew,;									//aHeader
    											aColsNew)									//aCols
			oMsNew:lInsert := .F.
			
		//Grupo Opções
		@ 183, 003 	GROUP oGrpOpc TO 220, (nJanLarg/2) 	PROMPT "Opções: " 	OF oDlgObs COLOR 0, 16777215 PIXEL
			//Chave
			@ 190, 005 SAY   oSayChave PROMPT "Campo Chave importado?"  SIZE 100, 011 OF oDlgObs COLORS 0, 16777215 PIXEL
			@ 200, 005 RADIO oRadChave VAR	nRadChave ITEMS "Conforme Arquivo","Conforme Sequencia do Protheus (SXE/SXF)" SIZE 120, 019 OF oDlgObs COLOR 0, 16777215 PIXEL
			oRadChave:lActive := lAtivChav
			
			//Filial
			@ 190, 180 SAY   oSayFilial PROMPT "Campo Filial importado?"  SIZE 100, 011 OF oDlgObs COLORS 0, 16777215 PIXEL
			@ 200, 180 RADIO oRadFilial VAR	nRadFilial ITEMS "Conforme Arquivo","Conforme Filial do Protheus (xFilial)" SIZE 120, 019 OF oDlgObs COLOR 0, 16777215 PIXEL
			
		//Grupo Ações
		@ 223, 003 	GROUP oGrpAco TO 247, (nJanLarg/2) 	PROMPT "Ações: " 		OF oDlgObs COLOR 0, 16777215 PIXEL
			@ 229, (nJanLarg/2)-(63*1)  BUTTON oBtnCanc PROMPT "Cancelar"      SIZE 60, 014 OF oDlgObs ACTION (lRet := .F., oDlgObs:End()) PIXEL
			@ 229, (nJanLarg/2)-(63*2)  BUTTON oBtnConf PROMPT "Confirmar"     SIZE 60, 014 OF oDlgObs ACTION (aHeadImp := aClone(oMsNew:aCols), lRet := .T., oDlgObs:End()) PIXEL
	ACTIVATE MSDIALOG oDlgObs CENTERED
	
	//Se a tela for confirmada, atualiza as variáveis
	If lRet
		lChvProt := nRadChave  == 2
		lFilProt := nRadFilial == 2
	EndIf
Return lRet
/*/{Protheus.doc} HVSPTA01
Validação do campo Tipo na tela de observação da carga genérica
@type function
@author Atilio
@since 25/10/2015
@version 1.0
@return lRetorn, Retorno da rotina
	@example
	u_HVSPTA01()
/*/
User Function HVSPTA01()
	Local lRetorn := .F.
	Local aColsAux := oMsNew:aCols
	Local nLinAtu := oMsNew:nAt
	
	//Se o campo atual estiver contido nos campos próprios do execauto (como LINPOS)
	If aColsAux[nLinAtu][01] $ cCampTipo
		lRetorn := .T.
	
	//Senão, campo não pode ser alterado
	Else
		lRetorn := .F.
		MsgAlert("Campo não pode ser alterado!", "Atenção")
	EndIf
	
Return lRetorn
/*---------------------------------------------------------------------------------*
 | Func.: fImport                                                                  |
 | Autor: Daniel Atilio                                                            |
 | Data:  25/10/2015                                                               |
 | Desc:  Função responsável por fazer a importação dos dados                      |
 *---------------------------------------------------------------------------------*/
Static Function fImport()
	Local nLinAtu	:= 2
	Local cLinAtu	:= ""
	Local aAuxAtu	:= {}
	Local cArqLog	:= ""
	Local cConLog	:= ""
	Local cSequen	:= StrZero(1, 10)
	Local nPosAux	:= 1
	Local lFalhou	:= .F.
	Local xConteud:= ""
	Local nAuxLog	:= 0
	Local aLogAuto:= {}
	Private aDados	:= {}
	Private aSC5 := {}
	Private aItem := {}
	Private aSC6 := {}
	Private lProc := .T.
	ProcRegua(nTotalReg)
	
	//Percorrendo os registros
	While !Ft_FEoF() .And. nLinAtu <= FT_FLastRec()
		IncProc("Analisando linha "+cValToChar(nLinAtu)+" de "+cValToChar(nTotalReg)+"...")
		cArqLog := "log_"+cCmbTip+"_lin_"+cValToChar(nLinAtu)+"_"+dToS(dDataBase)+"_"+StrTran(Time(), ":", "-")+".txt"
		cConLog := "Tipo:     "+cCmbTip+STR_PULA
		cConLog += "Usuário:  "+UsrRetName(RetCodUsr())+STR_PULA
		cConLog += "Ambiente: "+GetEnvServer()+STR_PULA
		cConLog += "Data:     "+dToC(dDataBase)+STR_PULA
		cConLog += "Hora:     "+Time()+STR_PULA
		cConLog += "----"+STR_PULA+STR_PULA
		
		//Pegando a linha atual e transformando em array
		cLinAtu := Ft_FReadLn()
		cLinAtu := Iif(SubStr(cLinAtu, Len(cLinAtu), 1) == ";", SubStr(cLinAtu, 1, Len(cLinAtu)-1), cLinAtu)
		aAuxAtu := Separa(cLinAtu, cGetCar)
		
		//Se tiver dados
		If !Empty(cLinAtu)
			//Se o tamanho for diferente, registra erro
			If Len(aAuxAtu) != Len(aHeadImp)
				cConLog += "O tamanho de campos da linha, difere do tamanho de campos do cabeçalho!"+STR_PULA
				cConLog += "Linha:     "+cValToChar(Len(aAuxAtu))+STR_PULA
				cConLog += "Cabeçalho: "+cValToChar(Len(aHeadImp))+STR_PULA
			
				//Gerando o arquivo
				MemoWrite(cDirTmp+cArqLog, cConLog)
			
				//Gravando o registro
				RecLock(cAliasTmp, .T.)
					TMP_SEQ	:= cSequen
					TMP_LINHA	:= nLinAtu
					TMP_ARQ	:= cArqLog
				(cAliasTmp)->(MsUnlock())
				
				//Incrementa a sequencia
				cSequen := Soma1(cSequen)

				IF EXISTBLOCK("CARGENP3")
					IF EXECBLOCK("CARGENP3",.F.,.F.)
						MsgStop(" CARGENP3 - Processo abortado")
						Exit
					Endif
				ENDIF
				
			//Senão, carrega as informações no array
			Else
				aDados	:= {}

				lFalhou:= .F.
				
				//Iniciando a transação
				Begin Transaction
					//Percorre o cabeçalho
					For nPosAux := 1 To Len(aHeadImp)
						xConteud := aAuxAtu[nPosAux]
						
						//Se o tipo do campo for Numérico
						If aHeadImp[nPosAux][2] == 'N'
							xConteud := Val(aAuxAtu[nPosAux])
						
						//Se o tipo for Lógico
						ElseIf aHeadImp[nPosAux][2] == 'L'
							xConteud := Iif(aAuxAtu[nPosAux] == '.T.', .T., .F.)
							
						//Se o tipo for Data
						ElseIf aHeadImp[nPosAux][2] == 'D'
							//Se tiver '/' na data, é padrão DD/MM/YYYY
							If '/' $ aAuxAtu[nPosAux]
								xConteud := cToD(aAuxAtu[nPosAux])
								
							//Senão, é padrão YYYYMMDD
							Else
								xConteud := sToD(aAuxAtu[nPosAux])
							EndIf
						ElseIf aHeadImp[nPosAux][2] == 'C'
							//Trunca o tamanho do campo conforme o tamanho da SX3 -- Adicionado por Raphael 02/06/2024
							If !('_FILIAL' $ aHeadImp[nPosAux][1]) .and. !('A1_CGC' $ aHeadImp[nPosAux][1])
								xConteud := SubStr(xConteud,1,TamSx3(aHeadImp[nPosAux][1])[1])
							Endif
						EndIf
						
						//Se for o campo filial
						If '_FILIAL' $ aHeadImp[nPosAux][1]
							//Se a filial for conforme o protheus
							If lFilProt
								xConteud := FWxFilial(cTabela)
							EndIf
						EndIf

						IF EXISTBLOCK("CARGENP1")
							xConteud := EXECBLOCK("CARGENP1",.F.,.F.,{aHeadImp,nPosAux,xConteud,aAuxAtu})
						ENDIF
						
						//Se for o campo chave
						If Alltrim(cCampoChv) == Alltrim(aHeadImp[nPosAux][1])
							//Se a chave for conforme o protheus
							If lChvProt
								xConteud := GetSXENum(cTabela, cCampoChv)
							EndIf
						EndIf
						
                        //Adicionando no vetor que será importado
                        aAdd(aDados,{	aHeadImp[nPosAux][1],;			//Campo
                                        xConteud,;							//Conteúdo
                                        Nil})								//Compatibilidade
                    
					Next

					IF EXISTBLOCK("CARGENP2")
						lProc := EXECBLOCK("CARGENP2",.F.,.F.,{aHeadImp,nPosAux,xConteud,aAuxAtu})
					ENDIF

					IF lProc
						//Seta as variáveis de log do protheus
						lAutoErrNoFile	:= .T.
						lMsErroAuto		:= .F.
						
						//Executa o execauto
						&(cRotina)
						
						//Se tiver alguma falha
						If lMsErroAuto
							lFalhou := .T.
							
							//Pegando log do ExecAuto
							aLogAuto := GetAutoGRLog()
		
							//Gerando log
							For nAuxLog :=1 To Len(aLogAuto)
								cConLog += aLogAuto[nAuxLog] + STR_PULA
							Next
							
							DisarmTransaction()
						EndIf
					Endif
				End Transaction
				
				//Se houve falha na importação, grava na tabela temporária
				If lFalhou
					//Gerando o arquivo
					MemoWrite(cDirTmp+cArqLog, cConLog)
				
					//Gravando o registro
					RecLock(cAliasTmp, .T.)
						TMP_SEQ	:= cSequen
						TMP_LINHA	:= nLinAtu
						TMP_ARQ	:= cArqLog
					(cAliasTmp)->(MsUnlock())
					
					//Incrementa a sequencia
					cSequen := Soma1(cSequen)
				EndIf
			EndIf
		EndIf
		
		nLinAtu++
		Ft_FSkip()
	EndDo
Return
/*---------------------------------------------------------------------------------*
 | Func.: fTelaErro                                                                |
 | Autor: Daniel Atilio                                                            |
 | Data:  25/10/2015                                                               |
 | Desc:  Função que mostra os erros gerados na tela                               |
 *---------------------------------------------------------------------------------*/
Static Function fTelaErro()
	Local aArea		:= GetArea()
	Local oDlgErro
	Local oGrpErr
	Local oGrpAco
	Local oBtnFech
	Local oBtnVisu
	Local nJanLarErr	:= 600
	Local nJanAltErr	:= 400
	//Criando a Janela
	DEFINE MSDIALOG oDlgErro TITLE "Erros na Importação" FROM 000, 000  TO nJanAltErr, nJanLarErr COLORS 0, 16777215 PIXEL
		//Grupo Erros
		@ 003, 003 	GROUP oGrpErr TO (nJanAltErr/2)-28, (nJanLarErr/2) 	PROMPT "Erros: " 		OF oDlgErro COLOR 0, 16777215 PIXEL
			//Criando o MsSelect
			oBrowChk := MsSelect():New(	cAliasTmp,;												//cAlias
											"",;														//cCampo
											,;															//cCpo
											aCampos,;													//aCampos
											,;															//lInv
											,;															//cMar
											{010, 006, (nJanAltErr/2)-31, (nJanLarErr/2)-3},;	//aCord
											,;															//cTopFun
											,;															//cBotFun
											oDlgErro,;													//oWnd
											,;															//uPar11
											)															//aColors
			oBrowChk:oBrowse:lHasMark    := .F.
			oBrowChk:oBrowse:lCanAllmark := .F.
		
		//Grupo Ações
		@ (nJanAltErr/2)-25, 003 	GROUP oGrpAco TO (nJanAltErr/2)-3, (nJanLarErr/2) 	PROMPT "Ações: " 		OF oDlgErro COLOR 0, 16777215 PIXEL
		
			//Botões
			@ (nJanAltErr/2)-18, (nJanLarErr/2)-(63*1)  BUTTON oBtnFech PROMPT "Fechar"        SIZE 60, 014 OF oDlgErro ACTION (oDlgErro:End()) PIXEL
			@ (nJanAltErr/2)-18, (nJanLarErr/2)-(63*2)  BUTTON oBtnVisu PROMPT "Vis.Erro"      SIZE 60, 014 OF oDlgErro ACTION (fVisErro()) PIXEL
	ACTIVATE MSDIALOG oDlgErro CENTERED
	
	RestArea(aArea)
Return
/*---------------------------------------------------------------------------------*
 | Func.: fVisErro                                                                 |
 | Autor: Daniel Atilio                                                            |
 | Data:  25/10/2015                                                               |
 | Desc:  Função que visualiza o erro conforme registro posicionado                |
 *---------------------------------------------------------------------------------*/
Static Function fVisErro()
	Local nRet := 0
	Local cNomeArq := Alltrim((cAliasTmp)->TMP_ARQ)
	//Tentando abrir o objeto
	nRet := ShellExecute("open", cNomeArq, "", cDirTmp, 1)
	
	//Se houver algum erro
	If nRet <= 32
		MsgStop("Não foi possí­vel abrir o arquivo " +cDirTmp+cNomeArq+ "!", "Atenção")
	EndIf
Return

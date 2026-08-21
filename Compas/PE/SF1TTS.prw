User Function SF1TTS

	if SF1->(FieldPos("F1_XNATURE")) > 0
		RecLock("SF1",.F.)
		SF1->F1_XNATURE := xNature
		SF1->(MsUnlock())
	eNDIF

Return

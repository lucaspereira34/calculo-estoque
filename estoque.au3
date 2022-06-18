#include <Array.au3>
#include <String.au3>
#include <File.au3>
#include <Date.au3>
#include <Excel.au3>

; Arrays que receberão os dados extraídos dos relatórios do SAP CCS
   Local $aT1, $aT2, $aT3, $aT4, $aT5

; Lê cada um dos arquivos de texto dos relatórios extraídos do SAP CCS para um array específico
   _FileReadToArray(@ScriptDir & '/T1.txt', $aT1, $FRTA_NOCOUNT, '|')
   _FileReadToArray(@ScriptDir & '/T2.txt', $aT2, $FRTA_NOCOUNT, '|')
   _FileReadToArray(@ScriptDir & '/T3.txt', $aT3, $FRTA_NOCOUNT, '|')
   _FileReadToArray(@ScriptDir & '/T4.txt', $aT4, $FRTA_NOCOUNT, '|')
   _FileReadToArray(@ScriptDir & '/T5.txt', $aT5, $FRTA_NOCOUNT, '|')

; Tratar T2 - Retirar notas em AGUARDOC
   For $i = UBound($aT2) - 1 to 0 Step -1
	  If StringReplace($aT2[$i][2], ' ', '') = 'AGUARDOC' Then
		 _ArrayDelete($aT2, $i)
	  EndIf
   Next

; Array ESTOQUE

Local $aEstoque[UBound($aT1)+1][26]

   ; Cabeçalho
	  $aEstoque[0][0] = 'Nota'
	  $aEstoque[0][1] = 'Validação Prazo'
	  $aEstoque[0][2] = 'Etapa'
	  $aEstoque[0][3] = 'GrpCódigos'
	  $aEstoque[0][4] = 'Texto code med.'
	  $aEstoque[0][5] = 'Status'
	  $aEstoque[0][6] = 'Prazo Regulatório'
	  $aEstoque[0][7] = 'Prazo Executado'
	  $aEstoque[0][8] = 'Data de Criação'
	  $aEstoque[0][9] = 'Validação da Verificação'
	  $aEstoque[0][10] = 'Data Agendamento Verificação'
	  $aEstoque[0][11] = 'Data da Verificação'
	  $aEstoque[0][12] = 'Dt. Sol. Doc.'
	  $aEstoque[0][13] = 'Dt. Env. Cart.'
	  $aEstoque[0][14] = 'Dt. Ent. Doc.'
	  $aEstoque[0][15] = 'Dt. Pagam.'
	  $aEstoque[0][16] = 'Prazo Verificação'
	  $aEstoque[0][17] = 'Txt. code codif.'
	  $aEstoque[0][18] = 'Local'
	  $aEstoque[0][19] = 'Bairro'
	  $aEstoque[0][20] = 'Rua'
	  $aEstoque[0][21] = 'Montante Inden.'
	  $aEstoque[0][22] = 'Nome do parceiro'
	  $aEstoque[0][23] = 'CenTrab respon.'
	  $aEstoque[0][24] = 'Data da Nota'
	  $aEstoque[0][25] = 'Qtd. Equip.'

; T1

   For $i = 0 to UBound($aT1) - 1
	  ; Número da Nota (retira todos os caracteres que não sejam dígitos)
		 $aEstoque[$i+1][0] = StringRegExpReplace($aT1[$i][1], '[\D, \h, \v]', '')

	  ; Data de criação
		 $aEstoque[$i+1][8] = StringReplace($aT1[$i][6], '.', '/')

	  ; Local
		 $aEstoque[$i+1][18] = StringStripWS($aT1[$i][10], $STR_STRIPTRAILING)

	  ; Bairro
		 $aEstoque[$i+1][19] = StringStripWS($aT1[$i][11], $STR_STRIPTRAILING)

	  ; Rua
		 $aEstoque[$i+1][20] = StringStripWS($aT1[$i][12], $STR_STRIPTRAILING)

	  ; Montante Indenizado
		 $aEstoque[$i+1][21] = StringStripWS($aT1[$i][15], $STR_STRIPTRAILING)

	  ; Nome do Parceiro
		 $aEstoque[$i+1][22] = StringStripWS($aT1[$i][17], $STR_STRIPTRAILING)

	  ; Centro de Trabalho Responsável
		 $aEstoque[$i+1][23] = StringStripWS($aT1[$i][18], $STR_STRIPTRAILING)

	  ; Data da Nota
		 $aEstoque[$i+1][24] = StringReplace($aT1[$i][19], '.', '/')
	  Next

; T2

   For $i = 1 to UBound($aEstoque) - 1
	  For $j = 0 to UBound($aT2) - 1
		 If $aEstoque[$i][0] = StringRegExpReplace($aT2[$j][1], '[\D, \h, \v]', '') Then
			; GrpCódigos -> Medida da Nota quando foi realizada a extração
			   $aEstoque[$i][3] = $aT2[$j][2]

			; Etapa 'Pagamento', 'Resposta' ou 'Verificação' de acordo com a Medida da Nota
			   If $aEstoque[$i][3] = 'PAGREALI' or $aEstoque[$i][3] = 'PAGAUTOR' Then
				  $aEstoque[$i][2] = 'Pagamento'
			   ElseIf $aEstoque[$i][3] = 'RECLIMPR' or $aEstoque[$i][3] = 'RECLPROC' or $aEstoque[$i][3] = 'NOTINTER' or $aEstoque[$i][3] = 'COMUNCLT' or $aEstoque[$i][3] = 'PARATEND' or $aEstoque[$i][3] = 'SOLICDOC' Then
				  $aEstoque[$i][2] = 'Resposta'
			   ElseIf $aEstoque[$i][3] = 'ANATECNI' or $aEstoque[$i][3] = 'VISTPROG' or $aEstoque[$i][3] = 'EANALISE' Then
				  $aEstoque[$i][2] = 'Verificação'
			   EndIf

			; Texto code med. (Retira todos os espaços extras)
			   $aEstoque[$i][4] = StringStripWS($aT2[$j][3], $STR_STRIPTRAILING)

			; Txt. code codif. (Retira todos os espaços extras)
			   $aEstoque[$i][17] = StringStripWS($aT2[$j][14], $STR_STRIPTRAILING)

			; Prazo Regulatório
			   If $aEstoque[$i][2] = 'Pagamento' Then
				  $aEstoque[$i][6] = 20
			   ElseIf $aEstoque[$i][2] = 'Resposta' or $aEstoque[$i][3] = 'ANATECNI' Then
				  $aEstoque[$i][6] = 15
			   ElseIf $aEstoque[$i][3] = 'VISTPROG' or 'EANALISE' Then
				  If StringLower(StringRight($aEstoque[$i][17], 7)) = 'urgente' Then
					 $aEstoque[$i][6] = 1
				  Else
					 $aEstoque[$i][6] = 10
				  EndIf
			   EndIf

			ExitLoop

		 EndIf
	  Next
   Next

; T4

   For $i = 1 to UBound($aEstoque) - 1
	  For $j = 0 to UBound($aT4) - 1
		 If $aEstoque[$i][0] = StringRegExpReplace($aT4[$j][1], '[\D, \h, \v]', '') and StringReplace($aT4[$j][21], ' ', '') <> '' Then
			; Prazo Verificação
			   If Int($aT4[$j][26]) <= Int($aT4[$j][25]) Then
				  $aEstoque[$i][16] = 'Dentro'
			   Elseif Int($aT4[$j][26]) > Int($aT4[$j][25]) Then
				  $aEstoque[$i][16] = 'Fora'
			   EndIf

			ExitLoop

		 EndIf
	  Next
   Next

; T3

   For $i = 1 to UBound($aEstoque) - 1
	  For $j = 0 to UBound($aT3) - 1
		 If $aEstoque[$i][0] = StringTrimLeft(StringRegExpReplace($aT3[$j][1], '[\D, \h, \v]', ''), 2) Then ; Remove todos os caracteres que não sejam digitos e remove os dois primeiros digitos (vem '00' da T3)
			; Data Agendamento Verificação
			   $aEstoque[$i][10] = StringReplace(StringReplace($aT3[$j][2], '.', '/'), ' ', '')

			; Data da Verificação
			   $aEstoque[$i][11] = StringReplace(StringReplace($aT3[$j][3], '.', '/'), ' ', '')

			; Data da Solicitação de Documentos
			   $aEstoque[$i][12] = StringReplace(StringReplace($aT3[$j][4], '.', '/'), ' ', '')

			; Data de Envio da Carta (Resposta)
			   $aEstoque[$i][13] = StringReplace(StringReplace($aT3[$j][5], '.', '/'), ' ', '')

			; Data de Entrega da Documentação
			   $aEstoque[$i][14] = StringReplace(StringReplace($aT3[$j][6], '.', '/'), ' ', '')

			; Data de Pagamento
			   $aEstoque[$i][15] = StringReplace(StringReplace($aT3[$j][7], '.', '/'), ' ', '')

			; Validação da Verificação
			   If $aEstoque[$i][10] = $aEstoque[$i][11] Then
				  $aEstoque[$i][9] = 'VERDADEIRO'
			   Else
				  $aEstoque[$i][9] = 'FALSO'
			   EndIf

			; Prazo Executado

			   ; Cria uma variável para as datas usadas para cálculo de prazo, alterando o formato de DD/MM/YYYY para YYYY/MM/DD (necessário para cálculo da função _DateDiff()
				  Local $sDtCriacao = StringRight($aEstoque[$i][8], 4) & '/' & StringMid($aEstoque[$i][8], 4, 2) & '/' & StringLeft($aEstoque[$i][8], 2)
				  Local $sDtVerific = StringRight($aEstoque[$i][11], 4) & '/' & StringMid($aEstoque[$i][11], 4, 2) & '/' & StringLeft($aEstoque[$i][11], 2)
				  Local $sDtSolicdoc = StringRight($aEstoque[$i][12], 4) & '/' & StringMid($aEstoque[$i][12], 4, 2) & '/' & StringLeft($aEstoque[$i][12], 2)
				  Local $sDtEnvio = StringRight($aEstoque[$i][13], 4) & '/' & StringMid($aEstoque[$i][13], 4, 2) & '/' & StringLeft($aEstoque[$i][13], 2)
				  Local $sDtEntrega = StringRight($aEstoque[$i][14], 4) & '/' & StringMid($aEstoque[$i][14], 4, 2) & '/' & StringLeft($aEstoque[$i][14], 2)

			   ; PAGAUTOR E PAGREALI
				  If $aEstoque[$i][2] = 'Pagamento' Then
					 ; Se o cálculo Data de Solicitação de Doc - Data de Criação foi negativo:
						If _DateDiff('D', $sDtCriacao, $sDtSolicdoc) < 0 Then
						   ; Se o prazo de envio da resposta foi ultrapassado:
							  If _DateDiff('D', $sDtEntrega, $sDtEnvio) > 15 Then
								 $aEstoque[$i][7] = _DateDiff('D', $sDtEntrega, $sDtEnvio) + _DateDiff('D', $sDtEnvio, _NowCalcDate()) - 15
						   ; Se o prazo de envio da resposta NÃO foi ultrapassado:
							  Else
								 $aEstoque[$i][7] = _DateDiff('D', $sDtEnvio, _NowCalcDate())
							  EndIf
					 ; Se o cálculo Data de Solicitação de Doc - Data de Criação NÃO foi negativo:
						Else
						   ; Se a VISTORIA foi realizada FORA do prazo ou NÃO foi realizada
						   If $aEstoque[$i][16] = 'Fora' or $aEstoque[$i][16] = '' Then
							  ; Se o prazo de envio da resposta foi ultrapassado:
								 If (_DateDiff('D', $sDtCriacao, $sDtSolicdoc) + _DateDiff('D', $sDtEntrega, $sDtEnvio)) > 15 Then
									$aEstoque[$i][7] = _DateDiff('D', $sDtCriacao, $sDtSolicdoc) + _DateDiff('D', $sDtEntrega, $sDtEnvio) + _DateDiff('D', $sDtEnvio, _NowCalcDate()) - 15
							  ; Se o prazo de envio da resposta NÃO foi ultrapassado:
								 Else
									$aEstoque[$i][7] = _DateDiff('D', $sDtEnvio, _NowCalcDate())
								 EndIf
						; Se a VISTORIA foi realizada DENTRO do prazo
						   ElseIf $aEstoque[$i][16] = 'Dentro' Then
							  ; Se o prazo de envio da resposta foi ultrapassado:
								 If (_DateDiff('D', $sDtVerific, $sDtSolicdoc) + _DateDiff('D', $sDtEntrega, $sDtEnvio)) > 15 Then
									$aEstoque[$i][7] = _DateDiff('D', $sDtVerific, $sDtSolicdoc) + _DateDiff('D', $sDtEntrega, $sDtEnvio) + _DateDiff('D', $sDtEnvio, _NowCalcDate()) - 15
							  ; Se o prazo de envio da resposta NÃO foi ultrapassado:
								 Else
									$aEstoque[$i][7] = _DateDiff('D', $sDtEnvio, _NowCalcDate())
								 EndIf
						   EndIf
						EndIf

			   ; Notas em COMUNCLT
				  Elseif $aEstoque[$i][3] = 'COMUNCLT' Then
					 $aEstoque[$i][7] = _DateDiff('D', $sDtCriacao, _NowCalcDate())

			   ; Notas em NOTINTER
				  ElseIf $aEstoque[$i][3] = 'NOTINTER' Then
					 ; Se o cálculo Data de Solicitação de Doc - Data de Criação foi negativo:
						If _DateDiff('D', $sDtCriacao, $sDtSolicdoc) < 0 Then
						   $aEstoque[$i][7] = _DateDiff('D', $sDtEntrega, _NowCalcDate())
					 ; Se o cálculo Data de Solicitação de Doc - Data de Criação NÃO foi negativo:
						Else
						   ; Se a VISTORIA foi realizada FORA do prazo ou NÃO foi realizada:
							  If $aEstoque[$i][16] = 'Fora' or $aEstoque[$i][16] = '' Then
								 $aEstoque[$i][7] = _DateDiff('D', $sDtCriacao, $sDtSolicdoc)
						   ; Se a VISTORIA foi realizada DENTRO do prazo:
							  ElseIf $aEstoque[$i][16] = 'Dentro' Then
								 $aEstoque[$i][7] = _DateDiff('D', $sDtVerific, $sDtSolicdoc)
							  EndIf
						EndIf

			   ; Notas em PARATEND, RECLIMPR ou RECLPROC
				  ElseIf $aEstoque[$i][3] = 'PARATEND' or $aEstoque[$i][3] = 'RECLIMPR' or $aEstoque[$i][3] = 'RECLPROC' Then
					 ; Se o cálculo Data de Solicitação de Doc - Data de Criação foi negativo:
						If _DateDiff('D', $sDtCriacao, $sDtSolicdoc) < 0 Then
						   $aEstoque[$i][7] = _DateDiff('D', $sDtEntrega, _NowCalcDate())
					 ; Se o cálculo Data de Solicitação de Doc - Data de Criação NÃO foi negativo:
						Else
						; Se NÃO houve solicitação de documento:
						   If $sDtSolicdoc = '//' Then
							  $aEstoque[$i][7] = _DateDiff('D', $sDtCriacao, _NowCalcDate())
						; Se houve solicitação de documento:
						   Else
							  ; Se a VISTORIA foi realizada FORA do prazo ou NÃO foi realizada:
								 If $aEstoque[$i][16] = 'Fora' or $aEstoque[$i][16] = '' Then
									$aEstoque[$i][7] = _DateDiff('D', $sDtCriacao, $sDtSolicdoc) + _DateDiff('D', $sDtEntrega, _NowCalcDate())
							  ; Se a VISTORIA foi realizada DENTRO do prazo:
								 ElseIf $aEstoque[$i][16] = 'Dentro' Then
									$aEstoque[$i][7] = _DateDiff('D', $sDtVerific, $sDtSolicdoc) + _DateDiff('D', $sDtEntrega, _NowCalcDate())
								 EndIf
						   EndIf
						EndIf

			   ; Notas em SOLICDOC ou ANATECNI
				  ElseIf $aEstoque[$i][3] = 'SOLICDOC' or $aEstoque[$i][3] = 'ANATECNI' Then
					 ; Se a VISTORIA foi realizada FORA do prazo ou NÃO foi realizada:
						If $aEstoque[$i][16] = 'Fora' or $aEstoque[$i][16] = '' Then
						   $aEstoque[$i][7] = _DateDiff('D', $sDtCriacao, _NowCalcDate())
					 ; Se a VISTORIA foi realizada DENTRO do prazo:
						ElseIf $aEstoque[$i][16] = 'Dentro' Then
						   $aEstoque[$i][7] = _DateDiff('D', $sDtVerific, _NowCalcDate())
						EndIf

			   ; Notas em VISTPROG ou EANALISE
				  ElseIf $aEstoque[$i][3] = 'VISTPROG' or $aEstoque[$i][3] = 'EANALISE' Then
					 $aEstoque[$i][7] = _DateDiff('D', $sDtCriacao, _NowCalcDate())

			   ; FIM do cálculo de Prazo Executado
				  EndIf

			; Validação Prazo
			   If $aEstoque[$i][7] <= $aEstoque[$i][6] Then
				  $aEstoque[$i][1] = 'Dentro'
			   Else
				  $aEstoque[$i][1] = 'Fora'
			   EndIf

			ExitLoop

		 EndIf
	  Next
   Next

; T5 (Quantidade de Equipamentos)

   Local $iQtd_equipamentos = 0

   For $i = 1 to UBound($aEstoque) - 1
	  For $j = 0 to UBound($aT5) - 1
		 If $aEstoque[$i][0] = StringRegExpReplace($aT5[$j][1], '[\D, \h, \v]', '') Then
			$iQtd_equipamentos = $iQtd_equipamentos + 1
		 EndIf
	  Next
	  $aEstoque[$i][25] = $iQtd_equipamentos
	  $iQtd_equipamentos = 0
   Next

; Salva planilha no excel

   ; Cria Objetos Excel (e o abre) e Workbook
   Local $oExcel = _Excel_Open()
   Local $oWorkbook = _Excel_BookNew($oExcel)

   ; Escreve o Array de Estoque no Workbook criado
   _Excel_RangeWrite($oWorkbook, Default, $aEstoque, 'A1')

   ; Variáveis de data e hora para nome do arquivo
	  Local $sData = StringRight(_NowCalcDate(), 2) & '-' & StringMid(_NowCalcDate(), 6, 2)
	  Local $sHora = StringReplace(_NowTime(4), ':', 'h')

   ; Salva Workbook na área de trabalho
	  _Excel_BookSaveAs($oWorkbook, @DesktopDir & '\Estoque' & $sData & '_' & $sHora & '.xlsx', $xlWorkbookDefault, True)
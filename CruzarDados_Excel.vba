Option Explicit

' ==============================================================================
' SELETIVOS AUTO MATCHER - MÓDULO EXCEL (VBA)
' Descrição: Sistema de Cruzamento e Classificação Regionalizada com Cotas
' ==============================================================================

Sub ProcessarClassificacaoRegionalizada()
    ' ----------------------------------------------------
    ' ⚙️ CONFIGURAÇÕES (Altere os nomes se precisar)
    ' ----------------------------------------------------
    Dim nomeAbaCand As String: nomeAbaCand = "Candidatos"
    Dim nomeAbaVagas As String: nomeAbaVagas = "Quadro de Vagas"
    
    Dim pctPCD As Double: pctPCD = 0.1 ' 10%
    Dim pctNegros As Double: pctNegros = 0.2 ' 20%
    ' ----------------------------------------------------
    
    Dim wsCand As Worksheet, wsVagas As Worksheet
    Dim totalLinhas As Long, totalCols As Long
    
    ' 1. Verifica se as abas existem
    On Error Resume Next
    Set wsCand = ThisWorkbook.Sheets(nomeAbaCand)
    Set wsVagas = ThisWorkbook.Sheets(nomeAbaVagas)
    On Error GoTo 0
    
    If wsCand Is Nothing Or wsVagas Is Nothing Then
        MsgBox "Erro: As abas '" & nomeAbaCand & "' e '" & nomeAbaVagas & "' precisam existir.", vbCritical
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    ' ----------------------------------------------------
    ' PASSO 1: Mapear o Quadro de Vagas na Memória
    ' ----------------------------------------------------
    Dim dictVagas As Object
    Set dictVagas = CreateObject("Scripting.Dictionary")
    dictVagas.CompareMode = vbTextCompare
    
    Dim arrVagas As Variant
    arrVagas = wsVagas.UsedRange.Value
    
    Dim vDRE As Long, vMun As Long, vCargo As Long, vQtd As Long
    vDRE = EncontrarColuna(arrVagas, "DRE")
    vMun = EncontrarColuna(arrVagas, "MUNICÍPIO")
    vCargo = EncontrarColuna(arrVagas, "CARGO")
    vQtd = EncontrarColuna(arrVagas, "VAGAS")
    
    Dim i As Long
    Dim chave As String
    For i = 2 To UBound(arrVagas, 1)
        chave = UCase(RemoverAcentos(arrVagas(i, vDRE)) & "_" & RemoverAcentos(arrVagas(i, vMun)) & "_" & RemoverAcentos(arrVagas(i, vCargo)))
        dictVagas(chave) = Val(arrVagas(i, vQtd))
    Next i
    
    ' ----------------------------------------------------
    ' PASSO 2: Ordenação em Cascata (O Desempate)
    ' ----------------------------------------------------
    Dim arrCand As Variant
    arrCand = wsCand.UsedRange.Value
    
    Dim cDRE As Long, cMun As Long, cCargo As Long
    Dim cNota As Long, cIdade As Long, cCE As Long, cCota As Long
    
    cDRE = EncontrarColuna(arrCand, "DRE")
    cMun = EncontrarColuna(arrCand, "MUNICÍPIO")
    cCargo = EncontrarColuna(arrCand, "CARGO")
    cNota = EncontrarColuna(arrCand, "NOTA FINAL")
    cIdade = EncontrarColuna(arrCand, "IDADE")
    cCE = EncontrarColuna(arrCand, "CE")
    cCota = EncontrarColuna(arrCand, "COTA")
    
    If cDRE = 0 Or cNota = 0 Or cCota = 0 Then
        MsgBox "Erro: Colunas essenciais (DRE, NOTA FINAL ou COTA) não encontradas.", vbCritical
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' Aciona o motor nativo do Excel para ordenar tudo
    With wsCand.Sort
        .SortFields.Clear
        .SortFields.Add Key:=wsCand.Columns(cDRE), Order:=xlAscending
        .SortFields.Add Key:=wsCand.Columns(cMun), Order:=xlAscending
        .SortFields.Add Key:=wsCand.Columns(cCargo), Order:=xlAscending
        .SortFields.Add Key:=wsCand.Columns(cNota), Order:=xlDescending ' Maior nota
        If cIdade > 0 Then .SortFields.Add Key:=wsCand.Columns(cIdade), Order:=xlDescending ' Mais velho
        If cCE > 0 Then .SortFields.Add Key:=wsCand.Columns(cCE), Order:=xlDescending ' Maior CE
        .SetRange wsCand.UsedRange
        .Header = xlYes
        .Apply
    End With
    
    ' ----------------------------------------------------
    ' PASSO 3: Agrupar e Aplicar as Cotas
    ' ----------------------------------------------------
    ' Recarrega o array agora que está ordenado
    arrCand = wsCand.UsedRange.Value
    totalLinhas = UBound(arrCand, 1)
    totalCols = UBound(arrCand, 2)
    
    Dim arrResult() As Variant
    ReDim arrResult(1 To totalLinhas, 1 To 5)
    
    ' Cabeçalhos novos
    arrResult(1, 1) = "CLASS GERAL"
    arrResult(1, 2) = "CLASS AC"
    arrResult(1, 3) = "CLASS NEGROS"
    arrResult(1, 4) = "CLASS PCD"
    arrResult(1, 5) = "BASE DE CÁLCULO"
    
    Dim r As Long, loteStart As Long, loteEnd As Long
    Dim atualGrupo As String, novoGrupo As String
    Dim vagas As Long, inscritos As Long, baseCalculo As Long
    Dim limPCD As Long, limNegros As Long
    Dim posGeral As Long, posAC As Long, posPCD As Long, posNegros As Long
    Dim strCota As String
    
    r = 2
    Do While r <= totalLinhas
        ' Identifica o Polo atual
        atualGrupo = UCase(RemoverAcentos(arrCand(r, cDRE)) & "_" & RemoverAcentos(arrCand(r, cMun)) & "_" & RemoverAcentos(arrCand(r, cCargo)))
        loteStart = r
        loteEnd = r
        
        ' Acha onde o grupo desse polo termina
        Do While loteEnd < totalLinhas
            novoGrupo = UCase(RemoverAcentos(arrCand(loteEnd + 1, cDRE)) & "_" & RemoverAcentos(arrCand(loteEnd + 1, cMun)) & "_" & RemoverAcentos(arrCand(loteEnd + 1, cCargo)))
            If novoGrupo = atualGrupo Then
                loteEnd = loteEnd + 1
            Else
                Exit Do
            End If
        Loop
        
        ' Cálculo dos limites do Lote
        inscritos = loteEnd - loteStart + 1
        vagas = 0
        If dictVagas.Exists(atualGrupo) Then vagas = dictVagas(atualGrupo)
        
        baseCalculo = IIf(vagas > 0, vagas, inscritos)
        limPCD = Application.WorksheetFunction.RoundUp(baseCalculo * pctPCD, 0)
        limNegros = Application.WorksheetFunction.RoundUp(baseCalculo * pctNegros, 0)
        
        posGeral = 1: posAC = 1: posPCD = 1: posNegros = 1
        
        ' Distribuição do Ranking
        For i = loteStart To loteEnd
            strCota = UCase(Trim(arrCand(i, cCota)))
            
            arrResult(i, 1) = posGeral
            posGeral = posGeral + 1
            arrResult(i, 2) = "-": arrResult(i, 3) = "-": arrResult(i, 4) = "-"
            
            If strCota = "PCD" Or strCota = "PESSOA COM DEFICIÊNCIA" Then
                arrResult(i, 4) = posPCD
                posPCD = posPCD + 1
            ElseIf strCota = "NEGROS" Or strCota = "COTA RACIAL" Then
                arrResult(i, 3) = posNegros
                posNegros = posNegros + 1
            Else
                arrResult(i, 2) = posAC
                posAC = posAC + 1
            End If
            
            arrResult(i, 5) = IIf(vagas > 0, vagas & " vagas", inscritos & " inscritos (CR)")
        Next i
        
        r = loteEnd + 1
    Loop
    
    ' ----------------------------------------------------
    ' PASSO 4: Descarregar Resultados na Planilha
    ' ----------------------------------------------------
    wsCand.Range(wsCand.Cells(1, totalCols + 1), wsCand.Cells(totalLinhas, totalCols + 5)).Value = arrResult
    wsCand.Range(wsCand.Cells(1, totalCols + 1), wsCand.Cells(1, totalCols + 5)).Interior.Color = RGB(226, 239, 218)
    
    Application.ScreenUpdating = True
    MsgBox "Classificação Regionalizada concluída!" & vbCrLf & "Agrupamento, desempate e cotas aplicados.", vbInformation, "Sucesso"
End Sub

' ==========================================
' FUNÇÕES AUXILIARES
' ==========================================
Function EncontrarColuna(arr As Variant, nomeProcurado As String) As Long
    Dim j As Long
    Dim limpoProcurado As String
    limpoProcurado = UCase(Replace(nomeProcurado, " ", ""))
    For j = 1 To UBound(arr, 2)
        If UCase(Replace(CStr(arr(1, j)), " ", "")) = limpoProcurado Then
            EncontrarColuna = j
            Exit Function
        End If
    Next j
    EncontrarColuna = 0
End Function

Function RemoverAcentos(ByVal texto As String) As String
    Dim i As Long
    Const comAcento As String = "ÁÀÃÂÄÉÈÊËÍÌÎÏÓÒÕÔÖÚÙÛÜÇáàãâäéèêëíìîïóòõôöúùûüç"
    Const semAcento As String = "AAAAAEEEEIIIIOOOOOUUUUCaaaaaeeeeiiiiooooouuuuc"
    For i = 1 To Len(comAcento)
        texto = Replace(texto, Mid(comAcento, i, 1), Mid(semAcento, i, 1))
    Next i
    RemoverAcentos = Trim(texto)
End Function

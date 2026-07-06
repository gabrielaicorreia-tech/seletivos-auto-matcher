Option Explicit

Sub CruzarDadosAutomatico()
    ' ==========================================
    ' ⚙️ CONFIGURAÇÕES GERAIS (Altere aqui)
    ' ==========================================
    Dim nomeAbaBase As String
    nomeAbaBase = "base"
    
    Dim colunasParaPuxar As Variant
    ' Liste o nome exato das colunas que deseja puxar da aba base
    colunasParaPuxar = Array("DATA FIM DO CONTRATO", "CH", "CB", "CE", "NOTAREDACAO", "NOTAFINAL", "CLASSGERAL", "CLASSAMPLA", "CLASSPCD", "CLASSRACA", "TIPO", "SELETIVO")
    ' ==========================================

    Dim wsBase As Worksheet, wsAtual As Worksheet
    Dim arrBase As Variant, arrAtual As Variant
    Dim dict As Object
    Dim i As Long, j As Long, c As Long
    
    ' 1. Verifica se a aba base existe
    On Error Resume Next
    Set wsBase = ThisWorkbook.Sheets(nomeAbaBase)
    On Error GoTo 0
    
    If wsBase Is Nothing Then
        MsgBox "Erro: Não encontrei a aba chamada exatamente '" & nomeAbaBase & "'.", vbCritical, "Erro de Execução"
        Exit Sub
    End If
    
    Set wsAtual = ActiveSheet
    If wsAtual.Name = wsBase.Name Then
        MsgBox "Erro: Você não pode rodar o cruzamento estando com a aba 'base' selecionada. Vá para a aba de destino.", vbExclamation, "Atenção"
        Exit Sub
    End If

    ' 2. Carrega os dados para a memória (Muito mais rápido que ler célula por célula)
    arrBase = wsBase.UsedRange.Value
    arrAtual = wsAtual.UsedRange.Value
    
    If IsEmpty(arrBase) Or IsEmpty(arrAtual) Then
        MsgBox "Erro: Uma das abas está vazia.", vbCritical, "Erro"
        Exit Sub
    End If

    ' 3. Localiza as colunas de CPF na Base e na Atual
    Dim baseIdxCPF As Long, atualIdxCPF As Long
    baseIdxCPF = EncontrarColuna(arrBase, "CPF")
    atualIdxCPF = EncontrarColuna(arrAtual, "CPF")
    
    If baseIdxCPF = 0 Or atualIdxCPF = 0 Then
        MsgBox "Erro: Coluna 'CPF' não encontrada na linha de cabeçalho.", vbCritical, "Erro"
        Exit Sub
    End If

    ' Localiza as colunas que vamos puxar da base
    Dim indicesBase() As Long
    ReDim indicesBase(LBound(colunasParaPuxar) To UBound(colunasParaPuxar))
    For c = LBound(colunasParaPuxar) To UBound(colunasParaPuxar)
        indicesBase(c) = EncontrarColuna(arrBase, CStr(colunasParaPuxar(c)))
    Next c

    ' 4. Cria o Dicionário de Busca na memória
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare ' Ignora maiúsculas/minúsculas
    
    Dim cpf As String
    Dim valoresExtraidos() As String
    
    For i = 2 To UBound(arrBase, 1)
        cpf = LimparCPF(CStr(arrBase(i, baseIdxCPF)))
        If cpf <> "" Then
            ReDim valoresExtraidos(LBound(indicesBase) To UBound(indicesBase))
            For c = LBound(indicesBase) To UBound(indicesBase)
                If indicesBase(c) > 0 Then
                    valoresExtraidos(c) = CStr(arrBase(i, indicesBase(c)))
                Else
                    valoresExtraidos(c) = ""
                End If
            Next c
            ' Salva no dicionário
            If Not dict.Exists(cpf) Then
                dict.Add cpf, valoresExtraidos
            End If
        End If
    Next i

    ' 5. Prepara o array final para a aba Atual
    Dim arrResultado() As Variant
    Dim totalLinhas As Long, totalColsAtual As Long, totalColsNovas As Long
    
    totalLinhas = UBound(arrAtual, 1)
    totalColsAtual = UBound(arrAtual, 2)
    totalColsNovas = UBound(colunasParaPuxar) - LBound(colunasParaPuxar) + 1
    
    ReDim arrResultado(1 To totalLinhas, 1 To totalColsAtual + totalColsNovas)
    
    Dim contSucesso As Long, contFalha As Long
    contSucesso = 0: contFalha = 0

    ' Preenche cabeçalhos
    For j = 1 To totalColsAtual
        arrResultado(1, j) = arrAtual(1, j)
    Next j
    For c = LBound(colunasParaPuxar) To UBound(colunasParaPuxar)
        arrResultado(1, totalColsAtual + c + 1) = colunasParaPuxar(c)
    Next c

    ' Cruza os dados
    Dim dadosPuxados As Variant
    For i = 2 To totalLinhas
        ' Copia linha original
        For j = 1 To totalColsAtual
            arrResultado(i, j) = arrAtual(i, j)
        Next j
        
        cpf = LimparCPF(CStr(arrAtual(i, atualIdxCPF)))
        
        If dict.Exists(cpf) Then
            dadosPuxados = dict(cpf)
            For c = LBound(dadosPuxados) To UBound(dadosPuxados)
                arrResultado(i, totalColsAtual + c + 1) = dadosPuxados(c)
            Next c
            contSucesso = contSucesso + 1
        Else
            For c = 1 To totalColsNovas
                arrResultado(i, totalColsAtual + c) = "Não encontrado"
            Next c
            contFalha = contFalha + 1
        End If
    Next i

    ' 6. Descarrega os dados na planilha de uma vez só
    wsAtual.Range(wsAtual.Cells(1, 1), wsAtual.Cells(totalLinhas, totalColsAtual + totalColsNovas)).Value = arrResultado
    
    ' Pinta o cabeçalho novo de verde
    wsAtual.Range(wsAtual.Cells(1, totalColsAtual + 1), wsAtual.Cells(1, totalColsAtual + totalColsNovas)).Interior.Color = RGB(217, 234, 211)

    MsgBox "Processamento concluído!" & vbCrLf & vbCrLf & _
           "✅ Localizados: " & contSucesso & vbCrLf & _
           "❌ Sem correspondência: " & contFalha, vbInformation, "Sucesso"
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

Function LimparCPF(texto As String) As String
    Dim i As Integer
    Dim apenasNumeros As String
    Dim charAtual As String
    
    ' Remove tudo que não for número
    For i = 1 To Len(texto)
        charAtual = Mid(texto, i, 1)
        If charAtual Like "[0-9]" Then
            apenasNumeros = apenasNumeros & charAtual
        End If
    Next i
    
    ' Adiciona zeros à esquerda até dar 11 dígitos
    If Len(apenasNumeros) > 0 Then
        LimparCPF = Right("00000000000" & apenasNumeros, 11)
    Else
        LimparCPF = ""
    End If
End Function

Attribute VB_Name = "ESTOQUE_P_RETIRADA"
Option Explicit

Sub ESTOQUE_P_RETIRADA()

    Dim wsE As Worksheet
    Dim wsR As Worksheet
    Dim ultimaLinha As Long
    Dim i As Long
    Dim col As Long
    Dim nome As String
    Dim dataMov As Date
    Dim temValor As Boolean
    Dim produto As String
    Dim qtd As Variant

    Set wsE = ThisWorkbook.Worksheets("Estoque")
    Set wsR = ThisWorkbook.Worksheets("Registro")

    ' =========================
    ' VERIFICAR NEGATIVOS
    ' =========================

    For i = 4 To 37

        If wsE.Cells(i, 8).Value <> "" Then

            If wsE.Cells(i, 8).Value < 0 Then
                MsgBox "Não é permitido número negativo!", vbCritical
                Exit Sub
            End If

            ' Verificar estoque suficiente
            If wsE.Cells(i, 8).Value > wsE.Cells(i, 3).Value Then
                MsgBox "Estoque insuficiente para o produto: " & _
                       wsE.Cells(i, 2).Value, vbCritical
                Exit Sub
            End If

        End If

    Next i

    ' =========================
    ' VERIFICAR SE HÁ ITENS
    ' =========================

    For i = 4 To 37

        If wsE.Cells(i, 8).Value <> "" Then
            temValor = True
            Exit For
        End If

    Next i

    If Not temValor Then
        MsgBox "Nenhum item informado!"
        Exit Sub
    End If

    ' =========================
    ' VERIFICAR NOME E DATA
    ' =========================

    If wsE.Range("H39").Value = "" Or wsE.Range("H40").Value = "" Then
        MsgBox "Preencha nome e data!"
        Exit Sub
    End If

    nome = wsE.Range("H39").Value
    dataMov = wsE.Range("H40").Value

    ' =========================
    ' REGISTRAR MOVIMENTAÇÃO
    ' =========================

    ultimaLinha = wsR.Cells(wsR.Rows.Count, 2).End(xlUp).Row + 1

    wsR.Cells(ultimaLinha, 2).Value = nome
    wsR.Cells(ultimaLinha, 3).Value = dataMov
    wsR.Cells(ultimaLinha, 4).Value = "RETIRADA"

    For i = 4 To 37

        produto = wsE.Cells(i, 2).Value
        qtd = wsE.Cells(i, 8).Value

        If qtd <> "" Then

            For col = 5 To wsR.Cells(2, wsR.Columns.Count).End(xlToLeft).Column

                If wsR.Cells(2, col).Value = produto Then
                    wsR.Cells(ultimaLinha, col).Value = qtd
                    Exit For
                End If

            Next col

        End If

    Next i

    ' =========================
    ' DESCONTAR DO ESTOQUE
    ' =========================

    For i = 4 To 37

        If wsE.Cells(i, 8).Value <> "" Then

            wsE.Cells(i, 3).Value = _
                wsE.Cells(i, 3).Value - wsE.Cells(i, 8).Value

        End If

    Next i

    ' =========================
    ' LIMPAR CAMPOS
    ' =========================

    wsE.Range("H4:H37").ClearContents
    wsE.Range("H39:H40").ClearContents

    MsgBox "Retirada registrada e estoque atualizado!"

End Sub

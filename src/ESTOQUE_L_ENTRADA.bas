Attribute VB_Name = "ESTOQUE_L_ENTRADA"
Option Explicit

Sub ESTOQUE_L_ENTRADA()

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

    Set wsE = ThisWorkbook.Worksheets("EstoqueLampada")
    Set wsR = ThisWorkbook.Worksheets("RegistroLampada")

    ' =========================
    ' VERIFICAR NEGATIVOS
    ' =========================

    For i = 4 To 13

        If wsE.Cells(i, 20).Value <> "" Then

            If wsE.Cells(i, 20).Value < 0 Then
                MsgBox "Não é permitido número negativo!", vbCritical
                Exit Sub
            End If

        End If

    Next i

    ' =========================
    ' VERIFICAR SE HÁ ITENS
    ' =========================

    For i = 4 To 13

        If wsE.Cells(i, 20).Value <> "" Then
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

    If wsE.Range("T15").Value = "" Or wsE.Range("T16").Value = "" Then
        MsgBox "Preencha nome e data!"
        Exit Sub
    End If

    nome = wsE.Range("T15").Value
    dataMov = wsE.Range("T16").Value

    ' =========================
    ' REGISTRAR MOVIMENTAÇÃO
    ' =========================

    ultimaLinha = wsR.Cells(wsR.Rows.Count, 3).End(xlUp).Row + 1

    wsR.Cells(ultimaLinha, 2).Value = nome
    wsR.Cells(ultimaLinha, 3).Value = dataMov
    wsR.Cells(ultimaLinha, 4).Value = "ENTRADA"

    For i = 4 To 13

        produto = wsE.Cells(i, 3).Value
        qtd = wsE.Cells(i, 20).Value

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
    ' ADICIONAR AO ESTOQUE
    ' =========================

    For i = 4 To 13

        If wsE.Cells(i, 20).Value <> "" Then

            wsE.Cells(i, 5).Value = _
                wsE.Cells(i, 5).Value + wsE.Cells(i, 20).Value

        End If

    Next i

    ' =========================
    ' LIMPAR CAMPOS
    ' =========================

    wsE.Range("T4:T13").ClearContents
    wsE.Range("T15:T16").ClearContents

    MsgBox "Entrada registrada e estoque atualizado!"

End Sub


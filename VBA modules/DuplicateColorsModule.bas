Attribute VB_Name = "DuplicateColorsModule"
Option Explicit

Public Sub DuplicateColors(rc As IRibbonControl)
    Dim continue As Integer, overfl As Boolean, i As Integer, k As Integer, n As Integer
    Dim dupes()
    ReDim dupes(1 To Selection.Cells.Count, 1 To 2)
    
    continue = MsgBox("������� ����� ������ � ��������� ������� �������" & vbCrLf & "(�������� 54 ������ ������). ����������?", vbYesNo, "DupliColor")
    If continue = 7 Then Exit Sub
    
    Selection.Interior.ColorIndex = -4142
    i = 3
    
    Dim currentCell As Variant
    For Each currentCell In Selection
        If WorksheetFunction.CountIf(Selection, currentCell.Value) > 1 Then
            For k = LBound(dupes) To UBound(dupes)
                If dupes(k, 1) = currentCell Then
                    currentCell.Interior.ColorIndex = dupes(k, 2)
                    n = n + 1
                End If
            Next k
            If currentCell.Interior.ColorIndex = -4142 Then
                If i = 57 Then
                    '���������� ������������� ���������� ������� ������ - ������ 54. �������� �������
                    i = 3
                    overfl = True
                End If
                currentCell.Interior.ColorIndex = i
                dupes(i, 1) = currentCell.Value
                dupes(i, 2) = i
                i = i + 1
            End If
        End If
    Next currentCell
    
    Dim result As String
    If overfl = True Then
        result = "����� 54"
    Else
        result = (i - 3)
    End If
    
    ' ����� ������ https://www.planetaexcel.ru/techniques/14/198/
    
    MsgBox "���������� ������� ������ " & result & vbCrLf & "���������� ������ " & n
    
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub aoc2ElfPRSGame1()
Dim Elf1 As String
Dim Elf2 As String
Dim count As Integer
For i = 1 To Range("A" & Rows.count).End(xlUp).Row
    Elf1 = Mid(Worksheets("Tabelle1").Cells(i, 1).Value, 1, 1)
    Elf2 = Mid(Worksheets("Tabelle1").Cells(i, 1).Value, 3, 1)
    If Elf1 = "A" And Elf2 = "Y" Then
    count = count + 8
    ElseIf Elf1 = "A" And Elf2 = "X" Then
    count = count + 4
    ElseIf Elf1 = "A" And Elf2 = "Z" Then
    count = count + 3
    ElseIf Elf1 = "B" And Elf2 = "Y" Then
    count = count + 5
    ElseIf Elf1 = "B" And Elf2 = "X" Then
    count = count + 1
    ElseIf Elf1 = "B" And Elf2 = "Z" Then
    count = count + 9
    ElseIf Elf1 = "C" And Elf2 = "Y" Then
    count = count + 2
    ElseIf Elf1 = "C" And Elf2 = "X" Then
    count = count + 7
    ElseIf Elf1 = "C" And Elf2 = "Z" Then
    count = count + 6
    End If
Next i
MsgBox "Points:" & count
End Sub

Sub aoc2ElfPRSGame2()
Dim Elf1 As String
Dim Elf2 As String
Dim count As Integer
For i = 1 To Range("A" & Rows.count).End(xlUp).Row
    Elf1 = Mid(Worksheets("Tabelle1").Cells(i, 1).Value, 1, 1)
    Elf2 = Mid(Worksheets("Tabelle1").Cells(i, 1).Value, 3, 1)
    If Elf1 = "A" And Elf2 = "Y" Then
    count = count + 4
    ElseIf Elf1 = "A" And Elf2 = "X" Then
    count = count + 3
    ElseIf Elf1 = "A" And Elf2 = "Z" Then
    count = count + 8
    ElseIf Elf1 = "B" And Elf2 = "Y" Then
    count = count + 5
    ElseIf Elf1 = "B" And Elf2 = "X" Then
    count = count + 1
    ElseIf Elf1 = "B" And Elf2 = "Z" Then
    count = count + 9
    ElseIf Elf1 = "C" And Elf2 = "Y" Then
    count = count + 6
    ElseIf Elf1 = "C" And Elf2 = "X" Then
    count = count + 2
    ElseIf Elf1 = "C" And Elf2 = "Z" Then
    count = count + 7
    End If
Next i
MsgBox "Points:" & count
End Sub

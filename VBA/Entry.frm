VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Entry 
   Caption         =   "Course Entry"
   ClientHeight    =   5040
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14805
   OleObjectBlob   =   "Entry.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Entry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub TextBox37_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim userDate As Date
    
    If IsDate(TextBox37.Value) Then
        userDate = CDate(TextBox37.Value)
        MsgBox "Valid date entered: " & userDate, vbInformation
    Else
        MsgBox "Please enter a valid date (e.g., MM/DD/YYYY).", vbExclamation
        TextBox37.SetFocus
    End If
End Sub

Private Sub addRound_Click()
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim l As Integer
    Dim m As Integer
    Dim emptyRow As Long
    Dim answer As VbMsgBoxResult
    
    Set ws = ThisWorkbook.Sheets("Score Database")
    Set tbl = ws.ListObjects("scoreDatabase")
    
    emptyRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row + 1
    
    answer = MsgBox("Do you want to add round to Database?", vbYesNo, "Confirm Round")
    
    If answer = vbYes Then
        ws.Cells(emptyRow, 1).Value = dateEntry.Value
        ws.Cells(emptyRow, 2).Value = courseEntry.Value
        
        For i = 1 To 18
            ws.Cells(emptyRow, i + 2).Value = Me.Controls("TextBox" & i).Value
        Next i
        
        For j = 1 To 18
            If Me.Controls("fairway" & j).Enabled = True Then
                If Me.Controls("fairway" & j).Value = True Then
                    ws.Cells(emptyRow, j + 38).Value = 1
                Else
                    ws.Cells(emptyRow, j + 38).Value = 0
                End If
            Else
                ws.Cells(emptyRow, j + 38).ClearContents
            End If
        Next j
        
        For k = 1 To 18
            If Me.Controls("green" & k).Value = True Then
                ws.Cells(emptyRow, k + 56).Value = 1
            Else
                ws.Cells(emptyRow, k + 56).Value = 0
            End If
        Next k
        
        For l = 1 To 18
            ws.Cells(emptyRow, l + 74).Value = Me.Controls("putts" & l).Value
        Next l
        
        For m = 1 To 9
            ws.Cells(emptyRow, m + 20).Value = Me.Controls("Label" & m + 21).Caption
        Next m
        
        For m = 10 To 18
            ws.Cells(emptyRow, m + 20).Value = Me.Controls("Label" & m + 22).Caption
        Next m
        
        If tbl.ListRows.Count > 0 Then
            ws.Cells(emptyRow, 120).Formula = "=SUM(D" & emptyRow & ":R" & emptyRow & ")"  ' Example formula for total score
        End If
    End If
End Sub



Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim countryRange As Range
    Dim cell As Range

    Set ws = ThisWorkbook.Sheets("Course Database")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    Set courseRange = ws.Range("A2:A" & lastRow)
    
    For Each cell In courseRange
            courseEntry.AddItem cell.Value
    Next cell
    
End Sub

Private Sub HandleTextBoxChange()
    Dim txt As MSForms.TextBox
    Set txt = Me.ActiveControl
    Dim totalSum As Integer
    Dim i As Integer
    Dim textBoxValue As Integer
    Dim textBoxName As String
    Dim totalBoxScore As Integer
    Dim frontNine As Integer
    
    totalSum = 0
    
    For i = 1 To 9
        textBoxName = "TextBox" & i
        If IsNumeric(Me.Controls(textBoxName).Value) Then
            textBoxValue = CInt(Me.Controls(textBoxName).Value)
        
        Else
            textBoxValue = 0
        End If
        
    frontNine = totalSum + textBoxValue
    frontScore.Caption = frontNine
    Next i
    
    For i = 10 To 18
        textBoxName = "TextBox" & i
        If IsNumeric(Me.Controls(textBoxName).Value) Then
            textBoxValue = CInt(Me.Controls(textBoxName).Value)
        
        Else
            textBoxValue = 0
        End If
        
    totalSum = totalSum + textBoxValue
    backNine = totalSum - frontNine
    backScore.Caption = backNine
    totalScore.Caption = totalSum
    Next i
    
End Sub
Private Sub courseEntry_Change()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim selectedCourse As String
    Dim i As Long
    Dim found As Boolean
    Dim LabelName As String
    Dim parValues(1 To 21) As Integer
    Dim courseRow As Long

    Set ws = ThisWorkbook.Sheets("Course Database")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).row
    
    selectedCourse = courseEntry.Value
    found = False
    
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = selectedCourse Then
            countryText.Caption = ws.Cells(i, 25).Value
            stateText.Caption = ws.Cells(i, 23).Value
            cityText.Caption = ws.Cells(i, 24).Value
            
            courseRow = i
            found = True
        End If
        Next i
    
    If found Then
        For j = 1 To 21
            parValues(j) = ws.Cells(courseRow, j + 1).Value
        Next j
        
        For j = 22 To 42
            LabelName = "Label" & j
            If Not IsError(parValues(j - 21)) Then
                Me.Controls(LabelName).Caption = CStr(parValues(j - 21))
            Else
                Me.Controls(LabelName).Caption = "Error"
            End If
        Next j
    Else
        countryLabel.Caption = ""
        stateLabel.Caption = ""
        cityLabel.Caption = ""
        
        For j = 22 To 42
            LabelName = "Label" & j
            Me.Controls(LabelName).Caption = ""
        Next j
    End If
    UpdateFairwayCheckboxes
End Sub
Private Sub CalculateScore()
    Dim totalSum As Integer
    Dim i As Integer
    Dim textBoxValue As Integer
    Dim textBoxName As String
    Dim totalBoxScore As Integer
    Dim frontNine As Integer
    
    totalSum = 0
    frontNine = 0
    For i = 1 To 9
        textBoxName = "TextBox" & i
        If IsNumeric(Me.Controls(textBoxName).Value) Then
            textBoxValue = CInt(Me.Controls(textBoxName).Value)
        
        Else
            textBoxValue = 0
        End If
        
    totalSum = totalSum + textBoxValue
    frontNine = totalSum
    frontScore.Caption = frontNine
    Next i
    
    For i = 10 To 18
        textBoxName = "TextBox" & i
        If IsNumeric(Me.Controls(textBoxName).Value) Then
            textBoxValue = CInt(Me.Controls(textBoxName).Value)
        
        Else
            textBoxValue = 0
        End If
        
    totalSum = totalSum + textBoxValue
    backNine = totalSum - frontNine
    backScore.Caption = backNine
    totalScore.Caption = totalSum
    Next i
    
End Sub
Private Sub AddPutts()
    Dim totalSum As Integer
    Dim i As Integer
    Dim textBoxValue As Integer
    Dim textBoxName As String
    Dim totalPutts As Integer
    Dim frontNine As Integer
    
    totalPutts = 0
    For i = 1 To 18
        textBoxName = "putts" & i
        If IsNumeric(Me.Controls(textBoxName).Value) Then
            textBoxValue = CInt(Me.Controls(textBoxName).Value)
            
        Else
            textBoxValue = 0
        End If
        
    totalPutts = totalPutts + textBoxValue
    puttsTotal.Caption = totalPutts
    Next i
End Sub
Sub CalculateCheckedFairways()
    Dim i As Integer
    Dim fairwaysChecked As Integer
    Dim checkboxName As String
    
    fairwaysChecked = 0
    For i = 1 To 18
        checkboxName = "fairway" & i
        
        If Me.Controls(checkboxName).Value = True Then
            fairwaysChecked = fairwaysChecked + 1
        End If
    Next i
    Me.totalFairways.Caption = fairwaysChecked
End Sub
Private Sub UpdateFairwayCheckboxes()
    Dim i As Integer
    
    For i = 1 To 9
        If Me.Controls("Label" & (21 + i)).Caption = "3" Then
            Me.Controls("fairway" & i).Enabled = False ' Disable checkbox for par 3
            Me.Controls("fairway" & i).Value = False
        Else
            Me.Controls("fairway" & i).Enabled = True ' Enable checkbox for non-par 3
        End If
        
        If Me.Controls("Label" & (31 + i)).Caption = "3" Then
            Me.Controls("fairway" & (i + 9)).Enabled = False ' Disable checkbox for par 3
            Me.Controls("fairway" & (i + 9)).Value = False
        Else
            Me.Controls("fairway" & (i + 9)).Enabled = True ' Enable checkbox for non-par 3
        End If
    Next i
End Sub
Sub CalculateCheckedGreens()
    Dim i As Integer
    Dim greensChecked As Integer
    Dim checkboxName As String
    
    greensChecked = 0
    For i = 1 To 18
        checkboxName = "green" & i
        If Me.Controls(checkboxName).Value = True Then
            greensChecked = greensChecked + 1
        End If
    Next i
    Me.totalGreens.Caption = greensChecked
End Sub
Private Sub green1_Click()
    CalculateCheckedGreens
End Sub

Private Sub green2_Click()
    CalculateCheckedGreens
End Sub

Private Sub green3_Click()
    CalculateCheckedGreens
End Sub

Private Sub green4_Click()
    CalculateCheckedGreens
End Sub

Private Sub green5_Click()
    CalculateCheckedGreens
End Sub

Private Sub green6_Click()
    CalculateCheckedGreens
End Sub

Private Sub green7_Click()
    CalculateCheckedGreens
End Sub

Private Sub green8_Click()
    CalculateCheckedGreens
End Sub

Private Sub green9_Click()
    CalculateCheckedGreens
End Sub

Private Sub green10_Click()
    CalculateCheckedGreens
End Sub

Private Sub green11_Click()
    CalculateCheckedGreens
End Sub

Private Sub green12_Click()
    CalculateCheckedGreens
End Sub

Private Sub green13_Click()
    CalculateCheckedGreens
End Sub

Private Sub green14_Click()
    CalculateCheckedGreens
End Sub

Private Sub green15_Click()
    CalculateCheckedGreens
End Sub

Private Sub green16_Click()
    CalculateCheckedGreens
End Sub

Private Sub green17_Click()
    CalculateCheckedGreens
End Sub

Private Sub green18_Click()
    CalculateCheckedGreens
End Sub
Private Sub Fairway1_Click()
    CalculateCheckedFairways
End Sub

Private Sub Fairway2_Click()
    CalculateCheckedFairways
End Sub

Private Sub Fairway3_Click()
    CalculateCheckedFairways
End Sub

Private Sub Fairway4_Click()
    CalculateCheckedFairways
End Sub

Private Sub Fairway5_Click()
    CalculateCheckedFairways
End Sub

Private Sub Fairway6_Click()
    CalculateCheckedFairways
End Sub

Private Sub Fairway7_Click()
    CalculateCheckedFairways
End Sub

Private Sub Fairway8_Click()
    CalculateCheckedFairways
End Sub

Private Sub Fairway9_Click()
    CalculateCheckedFairways
End Sub

Private Sub Fairway10_Click()
    CalculateCheckedFairways
End Sub

Private Sub Fairway11_Click()
    CalculateCheckedFairways
End Sub

Private Sub Fairway12_Click()
    CalculateCheckedFairways
End Sub

Private Sub Fairway13_Click()
    CalculateCheckedFairways
End Sub

Private Sub Fairway14_Click()
    CalculateCheckedFairways
End Sub

Private Sub Fairway15_Click()
    CalculateCheckedFairways
End Sub

Private Sub Fairway16_Click()
    CalculateCheckedFairways
End Sub

Private Sub Fairway17_Click()
    CalculateCheckedFairways
End Sub

Private Sub Fairway18_Click()
    CalculateCheckedFairways
End Sub

Private Sub TextBox1_Change()
    CalculateScore
End Sub

Private Sub TextBox2_Change()
    CalculateScore
End Sub

Private Sub TextBox3_Change()
    CalculateScore
End Sub

Private Sub TextBox4_Change()
    CalculateScore
End Sub

Private Sub TextBox5_Change()
    CalculateScore
End Sub

Private Sub TextBox6_Change()
    CalculateScore
End Sub

Private Sub TextBox7_Change()
    CalculateScore
End Sub

Private Sub TextBox8_Change()
    CalculateScore
End Sub

Private Sub TextBox9_Change()
    CalculateScore
End Sub

Private Sub TextBox10_Change()
    CalculateScore
End Sub

Private Sub TextBox11_Change()
    CalculateScore
End Sub

Private Sub TextBox12_Change()
    CalculateScore
End Sub

Private Sub TextBox13_Change()
    CalculateScore
End Sub

Private Sub TextBox14_Change()
    CalculateScore
End Sub

Private Sub TextBox15_Change()
    CalculateScore
End Sub

Private Sub TextBox16_Change()
    CalculateScore
End Sub

Private Sub TextBox17_Change()
    CalculateScore
End Sub

Private Sub TextBox18_Change()
    CalculateScore
End Sub
Private Sub putts1_Change()
    AddPutts
End Sub

Private Sub putts2_Change()
    AddPutts
End Sub

Private Sub putts3_Change()
    AddPutts
End Sub

Private Sub putts4_Change()
    AddPutts
End Sub

Private Sub putts5_Change()
    AddPutts
End Sub

Private Sub putts6_Change()
    AddPutts
End Sub

Private Sub putts7_Change()
    AddPutts
End Sub

Private Sub putts8_Change()
    AddPutts
End Sub

Private Sub putts9_Change()
    AddPutts
End Sub

Private Sub putts10_Change()
    AddPutts
End Sub

Private Sub putts11_Change()
    AddPutts
End Sub

Private Sub putts12_Change()
    AddPutts
End Sub

Private Sub putts13_Change()
    AddPutts
End Sub

Private Sub putts14_Change()
    AddPutts
End Sub

Private Sub putts15_Change()
    AddPutts
End Sub

Private Sub putts16_Change()
    AddPutts
End Sub

Private Sub putts17_Change()
    AddPutts
End Sub

Private Sub putts18_Change()
    AddPutts
End Sub
Private Sub fairway1_Enter()
    Me.fairway1.BackColor = RGB(0, 0, 0)
End Sub

Private Sub fairway1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.fairway1.BackColor = RGB(255, 255, 255)
End Sub

Private Sub fairway2_Enter()
    Me.fairway2.BackColor = RGB(0, 0, 0)
End Sub

Private Sub fairway2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.fairway2.BackColor = RGB(255, 255, 255)
End Sub

Private Sub fairway3_Enter()
    Me.fairway3.BackColor = RGB(0, 0, 0)
End Sub

Private Sub fairway3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.fairway3.BackColor = RGB(255, 255, 255)
End Sub

Private Sub fairway4_Enter()
    Me.fairway4.BackColor = RGB(0, 0, 0)
End Sub

Private Sub fairway4_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.fairway4.BackColor = RGB(255, 255, 255)
End Sub

Private Sub fairway5_Enter()
    Me.fairway5.BackColor = RGB(0, 0, 0)
End Sub

Private Sub fairway5_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.fairway5.BackColor = RGB(255, 255, 255)
End Sub

Private Sub fairway6_Enter()
    Me.fairway6.BackColor = RGB(0, 0, 0)
End Sub

Private Sub fairway6_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.fairway6.BackColor = RGB(255, 255, 255)
End Sub

Private Sub fairway7_Enter()
    Me.fairway7.BackColor = RGB(0, 0, 0)
End Sub

Private Sub fairway7_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.fairway7.BackColor = RGB(255, 255, 255)
End Sub

Private Sub fairway8_Enter()
    Me.fairway8.BackColor = RGB(0, 0, 0)
End Sub

Private Sub fairway8_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.fairway8.BackColor = RGB(255, 255, 255)
End Sub

Private Sub fairway9_Enter()
    Me.fairway9.BackColor = RGB(0, 0, 0)
End Sub

Private Sub fairway9_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.fairway9.BackColor = RGB(255, 255, 255)
End Sub

Private Sub fairway10_Enter()
    Me.fairway10.BackColor = RGB(0, 0, 0)
End Sub

Private Sub fairway10_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.fairway10.BackColor = RGB(255, 255, 255)
End Sub

Private Sub fairway11_Enter()
    Me.fairway11.BackColor = RGB(0, 0, 0)
End Sub

Private Sub fairway11_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.fairway11.BackColor = RGB(255, 255, 255)
End Sub

Private Sub fairway12_Enter()
    Me.fairway12.BackColor = RGB(0, 0, 0)
End Sub

Private Sub fairway12_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.fairway12.BackColor = RGB(255, 255, 255)
End Sub

Private Sub fairway13_Enter()
    Me.fairway13.BackColor = RGB(0, 0, 0)
End Sub

Private Sub fairway13_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.fairway13.BackColor = RGB(255, 255, 255)
End Sub

Private Sub fairway14_Enter()
    Me.fairway14.BackColor = RGB(0, 0, 0)
End Sub

Private Sub fairway14_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.fairway14.BackColor = RGB(255, 255, 255)
End Sub

Private Sub fairway15_Enter()
    Me.fairway15.BackColor = RGB(0, 0, 0)
End Sub

Private Sub fairway15_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.fairway15.BackColor = RGB(255, 255, 255)
End Sub

Private Sub fairway16_Enter()
    Me.fairway16.BackColor = RGB(0, 0, 0)
End Sub

Private Sub fairway16_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.fairway16.BackColor = RGB(255, 255, 255)
End Sub

Private Sub fairway17_Enter()
    Me.fairway17.BackColor = RGB(0, 0, 0)
End Sub

Private Sub fairway17_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.fairway17.BackColor = RGB(255, 255, 255)
End Sub

Private Sub fairway18_Enter()
    Me.fairway18.BackColor = RGB(0, 0, 0)
End Sub

Private Sub fairway18_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.fairway18.BackColor = RGB(255, 255, 255)
End Sub

Private Sub green1_Enter()
    Me.green1.BackColor = RGB(0, 0, 0)
End Sub

Private Sub green1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.green1.BackColor = RGB(255, 255, 255)
End Sub

Private Sub green2_Enter()
    Me.green2.BackColor = RGB(0, 0, 0)
End Sub

Private Sub green2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.green2.BackColor = RGB(255, 255, 255)
End Sub

Private Sub green3_Enter()
    Me.green3.BackColor = RGB(0, 0, 0)
End Sub

Private Sub green3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.green3.BackColor = RGB(255, 255, 255)
End Sub

Private Sub green4_Enter()
    Me.green4.BackColor = RGB(0, 0, 0)
End Sub

Private Sub green4_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.green4.BackColor = RGB(255, 255, 255)
End Sub

Private Sub green5_Enter()
    Me.green5.BackColor = RGB(0, 0, 0)
End Sub

Private Sub green5_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.green5.BackColor = RGB(255, 255, 255)
End Sub

Private Sub green6_Enter()
    Me.green6.BackColor = RGB(0, 0, 0)
End Sub

Private Sub green6_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.green6.BackColor = RGB(255, 255, 255)
End Sub

Private Sub green7_Enter()
    Me.green7.BackColor = RGB(0, 0, 0)
End Sub

Private Sub green7_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.green7.BackColor = RGB(255, 255, 255)
End Sub

Private Sub green8_Enter()
    Me.green8.BackColor = RGB(0, 0, 0)
End Sub

Private Sub green8_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.green8.BackColor = RGB(255, 255, 255)
End Sub

Private Sub green9_Enter()
    Me.green9.BackColor = RGB(0, 0, 0)
End Sub

Private Sub green9_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.green9.BackColor = RGB(255, 255, 255)
End Sub

Private Sub green10_Enter()
    Me.green10.BackColor = RGB(0, 0, 0)
End Sub

Private Sub green10_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.green10.BackColor = RGB(255, 255, 255)
End Sub

Private Sub green11_Enter()
    Me.green11.BackColor = RGB(0, 0, 0)
End Sub

Private Sub green11_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.green11.BackColor = RGB(255, 255, 255)
End Sub

Private Sub green12_Enter()
    Me.green12.BackColor = RGB(0, 0, 0)
End Sub

Private Sub green12_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.green12.BackColor = RGB(255, 255, 255)
End Sub

Private Sub green13_Enter()
    Me.green13.BackColor = RGB(0, 0, 0)
End Sub

Private Sub green13_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.green13.BackColor = RGB(255, 255, 255)
End Sub

Private Sub green14_Enter()
    Me.green14.BackColor = RGB(0, 0, 0)
End Sub

Private Sub green14_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.green14.BackColor = RGB(255, 255, 255)
End Sub

Private Sub green15_Enter()
    Me.green15.BackColor = RGB(0, 0, 0)
End Sub

Private Sub green15_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.green15.BackColor = RGB(255, 255, 255)
End Sub

Private Sub green16_Enter()
    Me.green16.BackColor = RGB(0, 0, 0)
End Sub

Private Sub green16_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.green16.BackColor = RGB(255, 255, 255)
End Sub

Private Sub green17_Enter()
    Me.green17.BackColor = RGB(0, 0, 0)
End Sub

Private Sub green17_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.green17.BackColor = RGB(255, 255, 255)
End Sub

Private Sub green18_Enter()
    Me.green18.BackColor = RGB(0, 0, 0)
End Sub

Private Sub green18_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.green18.BackColor = RGB(255, 255, 255)
End Sub
 

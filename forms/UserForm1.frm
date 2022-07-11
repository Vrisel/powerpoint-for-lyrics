VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "바로가기 버튼 생성"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   5040
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
    ReadListBox.Clear
    SelectedListBox.Clear
End Sub

Private Sub SelectCommandButton_Click()
    With SelectedListBox
        For i = 0 To ReadListBox.ListCount - 1
            If ReadListBox.Selected(i) Then
                    .AddItem (ReadListBox.List(i, 0))
                    If ReadListBox.List(i, 1) <> "" Then
                        .List(.ListCount - 1, 1) = ReadListBox.List(i, 1)
                    End If
                    ReadListBox.Selected(i) = False
            End If
        Next
        
        .height = 92.5
    End With
End Sub

Private Sub UpCommandButton_Click()
    With SelectedListBox
        If .ListIndex = 0 Then
            Exit Sub
        Else
            Dim tempList
            For i = 0 To 1
                tempList = .List(.ListIndex, i)
                .List(.ListIndex, i) = .List(.ListIndex - 1, i)
                .List(.ListIndex - 1, i) = tempList
            Next
            .ListIndex = .ListIndex - 1
        End If
    End With
End Sub
Private Sub DownCommandButton_Click()
    With SelectedListBox
        If .ListIndex = .ListCount - 1 Then
            Exit Sub
        Else
            Dim tempList
            For i = 0 To 1
                tempList = .List(.ListIndex, i)
                .List(.ListIndex, i) = .List(.ListIndex + 1, i)
                .List(.ListIndex + 1, i) = tempList
            Next
            .ListIndex = .ListIndex + 1
        End If
    End With
End Sub
Private Sub DelCommandButton_Click()
    'MultiSelect가 되는 경우
    'With SelectedListBox
    '    For i = 0 To .ListCount - 1
    '        If .Selected(i) Then
    '            .RemoveItem i
    '            Exit Sub
    '        End If
    '    Next
    'End With
    
    With SelectedListBox
        .RemoveItem (.ListIndex)
    End With
End Sub

Private Sub NextCommandButton_Click()
    If Me.SelectedListBox.ListCount < 1 Then
        'If MsgBox("아무 것도 선택되지 않았습니다. 종료하시겠습니까?", vbYesNo + vbDefaultButton2 + vbExclamation) = vbYes Then
        '    Unload Me
        'End If
        
        MsgBox "슬라이드를 선택해주세요.", vbOKOnly + vbExclamation
        Exit Sub
    End If
    
    Load UserForm2
    With UserForm2.SelectedListBox
        For i = 0 To Me.SelectedListBox.ListCount - 1
            .AddItem (Me.SelectedListBox.List(i, 0))
            .List(i, 1) = Me.SelectedListBox.List(i, 1)
            .List(i, 2) = recNotes(.List(i, 1))
        Next
        .Parent.StoredParam.Caption = Me.StoredParam.Caption
        
        Me.Hide
        .Parent.Show vbModal
    End With
    Unload Me
End Sub
Private Sub CancelCommandButton_Click()
    Unload Me
End Sub

Function recNotes(ByVal s As String)
    If InStr(1, s, "Verse") Then
        s = Replace(s, "Verse", "v")
    ElseIf InStr(1, s, "Pre-chorus") Then
        s = Replace(s, "Pre-chorus", "p")
    ElseIf InStr(1, s, "Chorus") Then
        s = Replace(s, "Chorus", "c")
    ElseIf InStr(1, s, "Bridge") Then
        s = Replace(s, "Bridge", "b")
    ElseIf InStr(1, s, "fin") Then
        s = Replace(s, "fin", "f")
    End If
    
    If InStr(1, s, "(") Then
    '    Dim sTemp() As String
    '    sTemp = Split(s, "(", 2)
    '    sTemp(1) = Split(sTemp(1), "/", 2)(0)
    '    s = sTemp(0) & sTemp(1)
    
        s = Split(s, "(", 2)(0)
    End If
    
    s = Replace(s, " ", "")
    
    recNotes = s
    Exit Function
End Function

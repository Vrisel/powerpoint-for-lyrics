VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 폼_바로가기버튼만들기1 
   Caption         =   "바로가기 버튼 생성"
   ClientHeight    =   3020
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   5040
   OleObjectBlob   =   "폼_바로가기버튼만들기1.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "폼_바로가기버튼만들기1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub UserForm_Initialize()
    리스트_슬라이드.Clear
    리스트_선택.Clear
End Sub

Private Sub 버튼_선택_Click()
    With 리스트_선택
        For i = 0 To 리스트_슬라이드.ListCount - 1
            If 리스트_슬라이드.Selected(i) Then
                    .AddItem (리스트_슬라이드.List(i, 0))
                    If 리스트_슬라이드.List(i, 1) <> "" Then
                        .List(.ListCount - 1, 1) = 리스트_슬라이드.List(i, 1)
                    End If
                    리스트_슬라이드.Selected(i) = False
            End If
        Next
        
        .height = 92.5
    End With
End Sub

Private Sub 버튼_위로_Click()
    With 리스트_선택
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
Private Sub 버튼_아래로_Click()
    With 리스트_선택
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
Private Sub 버튼_삭제_Click()
    'MultiSelect가 되는 경우
    'With 리스트_선택
    '    For i = 0 To .ListCount - 1
    '        If .Selected(i) Then
    '            .RemoveItem i
    '            Exit Sub
    '        End If
    '    Next
    'End With
    
    With 리스트_선택
        .RemoveItem (.ListIndex)
    End With
End Sub

Private Sub 버튼_다음_Click()
    If Me.리스트_선택.ListCount < 1 Then
        'If MsgBox("아무 것도 선택되지 않았습니다. 종료하시겠습니까?", vbYesNo + vbDefaultButton2 + vbExclamation) = vbYes Then
        '    Unload Me
        'End If
        
        MsgBox "슬라이드를 선택해주세요.", vbOKOnly + vbExclamation
        Exit Sub
    End If
    
    Load 폼_바로가기버튼만들기2
    With 폼_바로가기버튼만들기2
        With .리스트_선택
            For i = 0 To Me.리스트_선택.ListCount - 1
                .AddItem (Me.리스트_선택.List(i, 0))
                .List(i, 1) = Me.리스트_선택.List(i, 1)
                .List(i, 2) = 노트변환(.List(i, 1))
            Next
        End With
        .저장변수_대상구역인덱스.Caption = Me.저장변수_대상구역인덱스.Caption
        
        Me.Hide
        .Show vbModal
    End With
    Unload Me
End Sub
Private Sub 버튼_취소_Click()
    Unload Me
End Sub

Private Function 노트변환(ByVal s As String)
    '첫 단어 이니셜 추출
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
    
    '"(1/2)" 같은 문구가 있으면 제외하기 위해 "(" 뒤로는 삭제
    If InStr(1, s, "(") Then
        s = Split(s, "(", 2)(0)
    End If
    
    s = Replace(s, " ", "")
    
    노트변환 = s
    Exit Function
End Function

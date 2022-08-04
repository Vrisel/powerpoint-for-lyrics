VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 폼_바로가기버튼만들기1 
   Caption         =   "바로가기 버튼 생성"
   ClientHeight    =   3040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   5060
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

Private Sub 리스트_슬라이드_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    'With 리스트_슬라이드
        '.ListCount는 그저 아이템 개수라 원하는 기능이 아니고, _
            아래 멤버를 검색해서 찾아냈는데, 실제로 사용해보면 없는 멤버라고 나옴
        'If .ItemsSelected.Count > 0 Then
            슬라이드_선택
        'End If
    'End With
End Sub

Private Sub 버튼_선택_Click()
   슬라이드_선택
End Sub

Private Sub 슬라이드_선택()
    '선택된 항목들에 대해서만 반복문을 돌리고 싶은데, _
        버전 문제인지 관련 기능이 구현이 되어있지 않아 _
        모든 항목에 대해 선택여부를 확인하며 진행
    
    Dim i As Integer
    For i = 0 To 리스트_슬라이드.ListCount - 1
        If 리스트_슬라이드.Selected(i) Then
            With 리스트_선택
                .AddItem (리스트_슬라이드.List(i, 0))
                .List(.ListCount - 1, 1) = 리스트_슬라이드.List(i, 1)
            End With
            리스트_슬라이드.Selected(i) = False
        End If
    Next
End Sub

Private Sub 버튼_위로_Click()
    With 리스트_선택
        '이미 처음이거나 선택된 항목이 없을 때
        If .ListIndex = 0 Or .ListIndex < 0 Then
            Exit Sub
        
        Else
            '.ListIndex-1 자리에 항목 추가 -> .ListIndex가 1 늘어남, 새 항목과 index 2만큼 차이 남
            .AddItem .List(.ListIndex, 0), (.ListIndex - 1)
            
            .List((.ListIndex - 2), 1) = .List(.ListIndex, 1)
            .ListIndex = .ListIndex - 2
            .RemoveItem .ListIndex + 2
        End If
    End With
End Sub
Private Sub 버튼_아래로_Click()
    With 리스트_선택
        '이미 마지막이거나 선택된 항목이 없을 때
        If .ListIndex = .ListCount - 1 Or .ListIndex < 0 Then
            Exit Sub
        
        Else
            '.ListIndex+2 자리에 항목 추가 -> .ListIndex 변동 없음, 새 항목과 index 2만큼 차이 남
            .AddItem .List(.ListIndex, 0), (.ListIndex + 2)
            
            .List((.ListIndex + 2), 1) = .List(.ListIndex, 1)
            .ListIndex = .ListIndex + 2
            .RemoveItem .ListIndex - 2
        End If
    End With
End Sub
Private Sub 버튼_삭제_Click()
    With 리스트_선택
        'MultiSelect가 Single인 경우 _
            삭제 후 .ListIndex (혹은 .ListIndex-1)이 선택되어있는 편이 _
            다중삭제에 용이하기 때문에 _
            로직을 분리함
        If .MultiSelect = fmMultiSelectSingle Then
            If .ListIndex >= 0 Then
                .RemoveItem (.ListIndex)
            End If
            
        Else
            For i = .ListCount - 1 To 0 Step -1
                If .Selected(i) Then
                    .RemoveItem i
                End If
            Next
        End If
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
    If MsgBox("정말로 취소하시겠습니까?", vbOKCancel) = vbOK Then
        Unload Me
    End If
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

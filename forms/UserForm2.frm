VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "버튼 설명 수정"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    SelectedListBox.Clear
    ModifyTextBox.Text = ""
End Sub
Private Sub ModifyCommandButton_Click()
    With SelectedListBox
        If .ListIndex < 0 Then
            MsgBox "설명을 수정할 슬라이드를 선택해주세요.", vbOKOnly + vbExclamation
        Else
            .List(.ListIndex, 2) = ModifyTextBox.Text
            .ListIndex = -1
            ModifyTextBox.Text = ""
        End If
        
        .SetFocus
    End With
End Sub
Private Sub OKCommandButton_Click()
'구역 정보 갱신
    Dim secLBound As Long
    Dim secUBound As Long
    Dim secLen As Long
    
    secLBound = ActivePresentation.SectionProperties.FirstSlide(StoredParam.Caption)
    secUBound = ActivePresentation.SectionProperties.SlidesCount(StoredParam.Caption) + secLBound - 2
    secLen = secUBound - secLBound + 1
    
'버튼 개수
    Dim buttonCount As Integer
    buttonCount = SelectedListBox.ListCount
    
'index를 array로
    Dim indexArray() As String
    ReDim indexArray(buttonCount - 1)
    For i = 0 To UBound(indexArray)
        indexArray(i) = SelectedListBox.List(i, 0)
    Next
    
'String 생성
    Dim slideNote As String
    For i = 1 To buttonCount
        slideNote = slideNote & "[" & SelectedListBox.List(i - 1, 2) & "]"
        If i <> buttonCount Then
            slideNote = slideNote & " "
        End If
    Next
    
'확인 절차
    If MsgBox("해당 구역 (슬라이드 " & secLBound & "부터 " & secUBound & "까지)에" & vbCr _
            & "총 " & buttonCount & "개의 버튼을 생성하고" & vbCr _
            & "슬라이드 노트에 아래 문구를 추가하시겠습니까?" & vbCr _
            & ": " & slideNote, vbOKCancel + vbQuestion) = vbCancel Then
        Exit Sub
    End If
    
'범위 내 각 Slide에 대해 버튼과 노트 생성
    'Slides.Range에서 활용할 배열 생성
    Dim aSl() As Integer
    ReDim aSl(1 To secLen)
    For i = 1 To secLen
        aSl(i) = ActivePresentation.Slides(secLBound + i - 1).SlideIndex
    Next
    '초기화
    Dim iSwitch As Boolean
    If MsgBox("초기화 후 진행하시겠습니까?" & vbCr _
            & "(※주의: 텍스트 상자를 제외한 모든 도형이 사라집니다.)" _
            , vbYesNo + vbQuestion) = vbYes Then
        iSwitch = True
    Else
        iSwitch = False
    End If
    
'최종
    For Each oSl In ActivePresentation.Slides.Range(aSl)
        iCBWN iSwitch, oSl
        CreatButtons buttonCount, indexArray(), oSl
        EditNote slideNote, oSl
    Next
    
'끝
    MsgBox ("성공적으로 완료하였습니다.")
    Unload Me
End Sub
Private Sub CancelCommandButton_Click()
    Unload Me
End Sub

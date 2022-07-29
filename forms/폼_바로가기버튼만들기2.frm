VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 폼_바로가기버튼만들기2 
   Caption         =   "버튼 설명 수정"
   ClientHeight    =   3020
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   4560
   OleObjectBlob   =   "폼_바로가기버튼만들기2.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "폼_바로가기버튼만들기2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    리스트_선택.Clear
    텍스트_수정.text = ""
End Sub
Private Sub 버튼_수정_Click()
    With 리스트_선택
        If .ListIndex < 0 Then
            MsgBox "설명을 수정할 슬라이드를 선택해주세요.", vbOKOnly + vbExclamation
        Else
            .List(.ListIndex, 2) = 텍스트_수정.text
            .ListIndex = -1
            텍스트_수정.text = ""
        End If
        
        .SetFocus
    End With
End Sub
Private Sub 버튼_확인_Click()
    '구역 정보 갱신
    Dim 구역첫인덱스 As Long
    Dim 구역막인덱스 As Long
    Dim 구역길이 As Long
    
    구역첫인덱스 = ActivePresentation.SectionProperties.FirstSlide(저장변수_대상구역인덱스.Caption)
    구역막인덱스 = ActivePresentation.SectionProperties.SlidesCount(저장변수_대상구역인덱스.Caption) + 구역첫인덱스 - 2
    구역길이 = 구역막인덱스 - 구역첫인덱스 + 1
    
    '버튼 개수
    Dim 버튼개수 As Integer
    버튼개수 = 리스트_선택.ListCount
    
    'index를 array로
    Dim 바로갈인덱스배열() As String
    ReDim 바로갈인덱스배열(버튼개수 - 1)
    For i = 0 To UBound(바로갈인덱스배열)
        바로갈인덱스배열(i) = 리스트_선택.List(i, 0)
    Next
    
    '슬라이드 노트에 추가할 문구 생성
    Dim 슬라이드노트 As String
    For i = 0 To 버튼개수 - 1
        슬라이드노트 = 슬라이드노트 & "[" & 리스트_선택.List(i, 2) & "]"
        If i < 버튼개수 Then
            슬라이드노트 = 슬라이드노트 & " "
        End If
    Next
    
    '확인 절차
    If MsgBox("해당 구역 (슬라이드 " & 구역첫인덱스 & "부터 " & 구역막인덱스 & "까지)에" & vbCr _
            & "총 " & 버튼개수 & "개의 버튼을 생성하고" & vbCr _
            & "슬라이드 노트에 아래 문구를 추가하시겠습니까?" & vbCr _
            & ": " & 슬라이드노트, vbOKCancel + vbQuestion) = vbCancel Then
        Exit Sub
    End If
    
    'Slides.Range에서 활용할 배열 생성
    Dim 슬라이드배열() As Integer
    ReDim 슬라이드배열(구역길이)
    For i = 0 To 구역길이 - 1
        슬라이드배열(i) = ActivePresentation.Slides(구역첫인덱스 + i).SlideIndex
    Next
    
    버튼및노트삽입 슬라이드배열(), 바로갈인덱스배열(), 슬라이드노트
End Sub
Private Sub 버튼_취소_Click()
    If MsgBox("정말로 취소하시겠습니까?", vbOKCancel) = vbOK Then
        Unload Me
    End If
End Sub

Private Sub 버튼및노트삽입(슬라이드배열() As Integer, ByRef 바로갈인덱스배열() As String, 삽입문구 As String)
    Dim 초기화여부 As Boolean
    Select Case MsgBox("초기화 후 진행하시겠습니까?" & vbCr _
            & "(※주의: 텍스트 상자를 제외한 모든 도형이 사라집니다.)" _
            , vbYesNoCancel + vbQuestion)
    Case vbCancel
        Exit Sub
    Case vbYes
        초기화여부 = True
    Case vbNo
        초기화여부 = False
    End Select
    
    'For문 돌려서 버튼삽입 노트삽입
    Dim 슬라이드 As Slide
    For Each 슬라이드 In ActivePresentation.Slides.Range(슬라이드배열)
        버튼삽입 슬라이드, 초기화여부, 바로갈인덱스배열()
        슬라이드노트삽입 슬라이드, 초기화여부, 삽입문구
    Next
    
    MsgBox ("성공적으로 완료하였습니다.")
    Unload Me
End Sub

Private Sub 버튼삽입(슬라이드 As Slide, 초기화여부 As Boolean, ByRef 바로갈인덱스배열() As String)
    '초기화
    If 초기화여부 = True Then
        Dim i As Integer
        For i = 슬라이드.Shapes.Count To 1 Step -1
            With 슬라이드.Shapes(i)
                If .Type <> 14 And .Type <> 17 Then '14는 제목/본문, 17은 텍스트상자
                    .Delete
                End If
            End With
        Next i
    End If
    
    '버튼 생성
    Dim 버튼개수 As Integer
    버튼개수 = UBound(바로갈인덱스배열) - LBound(바로갈인덱스배열) + 1
    
    With 슬라이드.Parent.SlideMaster
        Dim 슬라이드w, 슬라이드h As Single
        슬라이드w = .width
        슬라이드h = .height
    End With
    
    Dim 버튼w, 버튼h, 버튼t As Single
    버튼w = 슬라이드w / 버튼개수
    버튼h = 슬라이드h / 2 + 30
    버튼t = 슬라이드h / 2
    
    For i = 0 To 버튼개수 - 1
        With 슬라이드.Shapes.AddShape(msoShapeRectangle, 버튼w * i, 버튼t, 버튼w, 버튼h)
            .Fill.Transparency = 1
            .Line.Visible = msoFalse
            .TextFrame2.VerticalAnchor = msoAnchorBottom
            .TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
            .TextFrame.TextRange.text = 바로갈인덱스배열(i)
            .ActionSettings(ppMouseClick).Action = ppActionRunMacro
            .ActionSettings(ppMouseClick).Run = "바로가기"
        End With
    Next
End Sub
Private Sub 슬라이드노트삽입(슬라이드 As Slide, 초기화여부 As Boolean, 삽입문구 As String)
    Dim 도형 As Shape
    '아래처럼 For Each - If로 밖에 접근할 수 없음. 인덱스3 으로는 접근 불가.
    For Each 도형 In 슬라이드.NotesPage.Shapes
        If 도형.PlaceholderFormat.Type = ppPlaceholderBody Then
            With 도형.TextFrame.TextRange
                If .Parent.HasText = True And 초기화여부 = True Then
                    .text = Split(.text, "[", 2)(0)
                    'RTrim 할 필요가..?
                End If
                
                If .Parent.HasText = True And Not Right(.text, 1) = vbCr Then
                    .InsertAfter vbCr
                End If
                
                .InsertAfter 삽입문구
            End With
            Exit Sub
        End If
    Next 도형
End Sub

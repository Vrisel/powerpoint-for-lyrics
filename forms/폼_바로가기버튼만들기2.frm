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
    텍스트_수정.Text = ""
End Sub
Private Sub 버튼_수정_Click()
    With 리스트_선택
        If .ListIndex < 0 Then
            MsgBox "설명을 수정할 슬라이드를 선택해주세요.", vbOKOnly + vbExclamation
        Else
            .List(.ListIndex, 2) = 텍스트_수정.Text
            .ListIndex = -1
            텍스트_수정.Text = ""
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
    
'String 생성
    Dim 슬라이드노트 As String
    For i = 1 To 버튼개수
        슬라이드노트 = 슬라이드노트 & "[" & 리스트_선택.List(i - 1, 2) & "]"
        If i <> 버튼개수 Then
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
    
'범위 내 각 Slide에 대해 버튼과 노트 생성
    'Slides.Range에서 활용할 배열 생성
    Dim 슬라이드배열() As Integer
    ReDim 슬라이드배열(1 To 구역길이)
    For i = 1 To 구역길이
        슬라이드배열(i) = ActivePresentation.Slides(구역첫인덱스 + i - 1).SlideIndex
    Next
    '초기화
    Dim 초기화여부 As Boolean
    If MsgBox("초기화 후 진행하시겠습니까?" & vbCr _
            & "(※주의: 텍스트 상자를 제외한 모든 도형이 사라집니다.)" _
            , vbYesNo + vbQuestion) = vbYes Then
        초기화여부 = True
    Else
        초기화여부 = False
    End If
    
'최종
    For Each 슬라이드 In ActivePresentation.Slides.Range(슬라이드배열)
        버튼생성전초기화 초기화여부, 슬라이드
        버튼생성 버튼개수, 바로갈인덱스배열(), 슬라이드
        슬라이드노트수정 슬라이드노트, 슬라이드
    Next
    
'끝
    MsgBox ("성공적으로 완료하였습니다.")
    Unload Me
End Sub
Private Sub 버튼_취소_Click()
    Unload Me
End Sub

Private Sub 버튼생성전초기화(ByVal 초기화여부 As Boolean, ByVal 슬라이드 As Slide)
    '(물어보고) 초기화
    '혹은 생성 직전에 상태를 파악해서 지우고 쓰겠냐고 되묻든가..
    '더 발전하면, 이미 있는 상태에서 수정..은 귀찮겠구나. 갯수 다르면 어쩔..
    If Not 초기화여부 Then
        Exit Sub
    End If
    
    Dim 도형 As Shape
DelAgain:
    For Each 도형 In 슬라이드.Shapes
        If 도형.Type <> 14 And 도형.Type <> 17 Then '14는 제목/본문, 17은 텍스트상자
            도형.Delete
        End If
    Next 도형
    
    For Each 도형 In 슬라이드.Shapes
        If 도형.Type <> 14 And 도형.Type <> 17 Then '14는 제목/본문, 17은 텍스트상자
            GoTo DelAgain
        End If
    Next 도형
    
    For Each 도형 In 슬라이드.NotesPage.Shapes
        If 도형.PlaceholderFormat.Type = ppPlaceholderBody Then
            With 도형.TextFrame.TextRange
                If .Parent.HasText Then
                    .Text = Split(.Text, "[", 2)(0) 'RTrim 할 필요가..?
                End If
            End With
        End If
    Next 도형
End Sub
Private Sub 버튼생성(ByVal n As Integer, ByRef 바로갈인덱스배열() As String, ByVal 슬라이드 As Slide)
    With 슬라이드
        Dim w As Single
        Dim h As Single
        w = .Parent.SlideMaster.width
        h = .Parent.SlideMaster.height
        For j = 0 To n - 1
            With .Shapes.AddShape(msoShapeRectangle, j * w / n, h / 2, w / n, h / 2 + 30)
                .Fill.Transparency = 1
                .Line.Visible = msoFalse
                .TextFrame2.VerticalAnchor = msoAnchorBottom
                .TextFrame.TextRange.Font.Color.RGB = RGB(0, 0, 0)
                .TextFrame.TextRange.Text = 바로갈인덱스배열(j)
                .ActionSettings(ppMouseClick).Action = ppActionRunMacro
                .ActionSettings(ppMouseClick).Run = "바로가기"
            End With
        Next
    End With
End Sub
Private Sub 슬라이드노트수정(ByVal 노트 As String, ByVal 슬라이드 As Slide)
    Dim 도형 As Shape
    '아래처럼 For Each - If로 밖에 접근할 수 없음. 인덱스3 으로는 접근 불가.
    For Each 도형 In 슬라이드.NotesPage.Shapes
        If 도형.PlaceholderFormat.Type = ppPlaceholderBody Then
            With 도형.TextFrame.TextRange
                If .Parent.HasText Then
                    If Not Right(.Text, 1) = vbCr Then
                        .InsertAfter vbCr
                    End If
                End If
                
                .InsertAfter 노트
            End With
            Exit Sub
            
        End If
    Next 도형
End Sub


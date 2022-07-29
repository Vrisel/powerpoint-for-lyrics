VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 폼_가사가져오기 
   Caption         =   "가사 입력"
   ClientHeight    =   3850
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   4560
   OleObjectBlob   =   "폼_가사가져오기.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "폼_가사가져오기"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    텍스트_제목.text = ""
    텍스트_가사.text = ""
End Sub

Private Sub 버튼_파일열기_Click()
    파일열기
End Sub
Private Sub 버튼_확인_Click()
    '텍스트 가공 작업
    Dim 가사분리() As String
    가사분리 = Split(텍스트_가사.text, vbCrLf & "//" & vbCrLf)
    
    With ActivePresentation.Slides
        '구역 생성을 위해 최초 삽입 인덱스 기억
        Dim 삽입위치 As Integer
        삽입위치 = .Count + 1
    
        '슬라이드 생성
        Dim 가사 As Variant
        For Each 가사 In 가사분리
            가사입력 가사
        Next 가사
        
        '마지막 빈 슬라이드 추가
        .AddSlide (.Count + 1), ActivePresentation.SlideMaster.CustomLayouts(2)
    End With
    
    '구역 생성 및 제목 입력
    ActivePresentation.SectionProperties.AddBeforeSlide 삽입위치, 텍스트_제목.text
    ActivePresentation.Slides(삽입위치).Shapes(1).TextFrame.TextRange.text = 텍스트_제목.text
    
    Unload Me
End Sub
Private Sub 버튼_취소_Click()
    Unload Me
End Sub

Private Sub 파일열기()
    Dim 제목, 가사 As String
    제목 = ""
    가사 = ""
    
    '파일 열기
    Dim 파일경로 As String
    With Application.FileDialog(msoFileDialogFilePicker)
        .Filters.Clear
        .Filters.Add "텍스트 파일", "*.txt"
        
        If .Show = True Then
            파일경로 = .SelectedItems(1)
            
            '정규식 사용은 도구(T) > 참조(R)... > Microsoft VBScript Regular Expression 5.5 추가
            With CreateObject("VBScript.RegExp")
                'VBA에서 백슬래시 이스케이프가 안 되어 Chr(92)로 대체
                '\\: 백슬래시 이스케이프
                '([^\\]+): 백슬래시가 아닌 문자 1개 이상, 캡처
                '\.txt$: ".txt"로 끝나는 확장자
                .Pattern = Chr(92) & Chr(92) & _
                            "([^" & Chr(92) & Chr(92) & "]+)" & _
                            Chr(92) & ".txt$"
                .IgnoreCase = True
                
                제목 = .Execute(파일경로).Item(0).SubMatches(0)
            End With
            
            With CreateObject("ADODB.Stream")
                .Open
                .Type = 2 'adTypeText
                .Charset = "UTF-8"
                .LoadFromFile = 파일경로
                .LineSeparator = 13 'adCR
                
                Do Until .EOS
                    가사 = 가사 & .ReadText(-2)
                Loop
            End With
    
            Me.텍스트_제목.text = 제목
            Me.텍스트_가사.text = 가사
        End If
    End With
End Sub

Private Sub 가사입력(가사 As Variant)
    Dim 가사분리() As String
    가사분리 = Split(가사, vbCrLf & "&&" & vbCrLf)
    
    '버튼 정보가 저장된 마지막줄의 경우
    If Left(가사분리(0), 1) = "[" Then
        '미구현
        
    '일반적인 경우
    Else
        With ActivePresentation.Slides
            '슬라이드 생성
            '레이아웃은 "제목 및 내용(#2)"
            .AddSlide (.Count + 1), ActivePresentation.SlideMaster.CustomLayouts(2)
            
            With .Item(.Count)
                '내용칸(#2)에 가사 입력
                .Shapes(2).TextFrame.TextRange.text = 가사분리(0)
                
                '슬라이드노트에 노트 입력 (있으면)
                If UBound(가사분리) = 1 Then
                    Dim 도형 As Shape
                    
                    '아래처럼 For Each - If로 밖에 접근할 수 없음. 인덱스3 으로는 접근 불가.
                    For Each 도형 In .NotesPage.Shapes
                        If 도형.PlaceholderFormat.Type = ppPlaceholderBody Then
                            도형.TextFrame.TextRange.text = 가사분리(1)
                            Exit Sub
                        End If
                    Next 도형
                End If
            End With
        End With
    End If
End Sub


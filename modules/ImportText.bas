Attribute VB_Name = "ImportText"
Sub 텍스트to슬라이드()
    '텍스트파일을 직접 읽어들이면 좋겠지만 아직 기능을 못 찾아서;;
    '그냥 UserForm에다가 때려박기... -_-;;;
    Load ufImportText
    With ufImportText
        .StoredParam.Caption = ""
        .Show vbModal
    End With
End Sub
Sub InputLyrics(oSl As Slide, s As String)
    Dim SplText() As String
    SplText() = Split(s, vbCrLf & "&&" & vbCrLf)
    
    '마지막이라서 특정 어쩌구인 경우
    If Left(SplText(0), 1) = "[" Then
        '
    '일반적인 경우
    Else
        '내용 입력
        oSl.Shapes(2).TextFrame.TextRange.Text = SplText(0)
        '노트 입력 (있으면)
        If UBound(SplText) = 1 Then
            Dim oSh As Shape
            For Each oSh In oSl.NotesPage.Shapes
                '아래 If가 왜 필요한지 모르겠지만 없으면 오류남.. 암튼 보통 인덱스3 인듯
                If oSh.PlaceholderFormat.Type = ppPlaceholderBody Then
                    oSh.TextFrame.TextRange.Text = SplText(1)
                    Exit Sub
                End If
            Next oSh
        End If
    End If
End Sub

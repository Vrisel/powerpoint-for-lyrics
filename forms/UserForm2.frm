VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "��ư ���� ����"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   470
   ClientWidth     =   4560
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  '������ ���
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
            MsgBox "������ ������ �����̵带 �������ּ���.", vbOKOnly + vbExclamation
        Else
            .List(.ListIndex, 2) = ModifyTextBox.Text
            .ListIndex = -1
            ModifyTextBox.Text = ""
        End If
        
        .SetFocus
    End With
End Sub
Private Sub OKCommandButton_Click()
'���� ���� ����
    Dim secLBound As Long
    Dim secUBound As Long
    Dim secLen As Long
    
    secLBound = ActivePresentation.SectionProperties.FirstSlide(StoredParam.Caption)
    secUBound = ActivePresentation.SectionProperties.SlidesCount(StoredParam.Caption) + secLBound - 2
    secLen = secUBound - secLBound + 1
    
'��ư ����
    Dim buttonCount As Integer
    buttonCount = SelectedListBox.ListCount
    
'index�� array��
    Dim indexArray() As String
    ReDim indexArray(buttonCount - 1)
    For i = 0 To UBound(indexArray)
        indexArray(i) = SelectedListBox.List(i, 0)
    Next
    
'String ����
    Dim slideNote As String
    For i = 1 To buttonCount
        slideNote = slideNote & "[" & SelectedListBox.List(i - 1, 2) & "]"
        If i <> buttonCount Then
            slideNote = slideNote & " "
        End If
    Next
    
'Ȯ�� ����
    If MsgBox("�ش� ���� (�����̵� " & secLBound & "���� " & secUBound & "����)��" & vbCr _
            & "�� " & buttonCount & "���� ��ư�� �����ϰ�" & vbCr _
            & "�����̵� ��Ʈ�� �Ʒ� ������ �߰��Ͻðڽ��ϱ�?" & vbCr _
            & ": " & slideNote, vbOKCancel + vbQuestion) = vbCancel Then
        Exit Sub
    End If
    
'���� �� �� Slide�� ���� ��ư�� ��Ʈ ����
    'Slides.Range���� Ȱ���� �迭 ����
    Dim aSl() As Integer
    ReDim aSl(1 To secLen)
    For i = 1 To secLen
        aSl(i) = ActivePresentation.Slides(secLBound + i - 1).SlideIndex
    Next
    '�ʱ�ȭ
    Dim iSwitch As Boolean
    If MsgBox("�ʱ�ȭ �� �����Ͻðڽ��ϱ�?" & vbCr _
            & "(������: �ؽ�Ʈ ���ڸ� ������ ��� ������ ������ϴ�.)" _
            , vbYesNo + vbQuestion) = vbYes Then
        iSwitch = True
    Else
        iSwitch = False
    End If
    
'����
    For Each oSl In ActivePresentation.Slides.Range(aSl)
        iCBWN iSwitch, oSl
        CreatButtons buttonCount, indexArray(), oSl
        EditNote slideNote, oSl
    Next
    
'��
    MsgBox ("���������� �Ϸ��Ͽ����ϴ�.")
    Unload Me
End Sub
Private Sub CancelCommandButton_Click()
    Unload Me
End Sub

Attribute VB_Name = "Module1"
Sub MailSystem()


Dim address, shA, shB, sign As Worksheet
Set address = ThisWorkbook.Sheets("���M��")
Set sign = ThisWorkbook.Sheets("����")

'�����������쐬
'Dim signcontents As String
'Dim signlastr As Integer
'signlastr = sign.Cells(Rows.Count, 2).End(xlUp).Row '�����V�[�g�̍ŏI�s���擾
'
'Call signcreate(signlastr)





'���M��ɐݒ肳��Ă���Ώێ҂��ォ�珇�ɏ���
Dim lastr As Integer
lastr = address.Cells(Rows.Count, 1).End(xlUp).Row

Dim r As Integer
Dim sheetname As String
For r = 2 To lastr  '�u���M��v1�s�ڂ���Ō�܂�
    Dim NAME, MA, body As String
    NAME = address.Cells(r, 1).Value
    MA = address.Cells(r, 2).Value
    sheetname = address.Cells(r, 3).Value
    
    
 '�����E�{���̐ݒ�
    '����
    Dim TITLE, BODIES As String
    TITLE = Sheets(sheetname).Cells(1, 2)

    '�{��
    Dim sentences() As String
    ReDim sentences(0)
    sentences(0) = Sheets(sheetname).Cells(2, 2)


    Dim sentencecounter, lastsentence As Integer
    lastsentence = Sheets(sheetname).Cells(Rows.Count, 2).End(xlUp).Row - 1
    
    For sentencecounter = 1 To lastsentence
        ReDim Preserve sentences(sentencecounter)
        sentences(sentencecounter) = Sheets(sheetname).Cells(sentencecounter + 2, 2)
    Next
   BODIES = Join(sentences, vbCrLf)
   
   MsgBox BODIES

  

    'Outlook���p�̏����i�A�v�����J���A�V�������[�����쐬�j
    Dim objOutlook As Outlook.Application
    Dim objmail As Outlook.MailItem

    Set objOutlook = New Outlook.Application
    Set objmail = objOutlook.CreateItem(olMailItem)

    With objmail
     .To = MA  '����=���M��V�[�g�̃��[���A�h���X
     .Subject = TITLE
     .body = BODIES
    End With
    
    objmail.Save
    
Next  '�u���M��v1�s�ڂ���Ō�܂Łi�I�j

End Sub

Sub signcreate(signlastr)

Dim signr(1 To signlastr) As String
Dim counter1 As Integer

Do Until counter1 = signlastr
    signr(counter1) = sign.Cells(counter, 2)
    counter1 = counter1 + 1
Loop

signcontents = Join(signr, vbCrLf)
End Sub

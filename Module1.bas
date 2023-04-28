Attribute VB_Name = "Module1"
Sub MailSystem()


Dim address, sign, searching_sheet As Worksheet
Set address = ThisWorkbook.Sheets("���M��")
Set sign = ThisWorkbook.Sheets("����")

Dim signature As String
'�����������쐬
signature = SetSentences(sign, 1)

'�{���̐ݒ肳�ꂽ�V�[�g�𑖍��A
'innder_dic�o'Subject':B1�̒l, 'Body':B2�`�̒l�����s�łȂ���������,�p���쐬
'inner_dic���V�[�g�����L�[�ɁAmail_dic�֊i�[
Dim mail_dic, body_dic As Object
Set mail_dic = CreateObject("scripting.dictionary")

'�u�b�N���V�[�g�𑖍�
Dim i, x, sheet_num, last_body_row As Long
Dim subject, body As String
sheet_num = ThisWorkbook.Sheets.Count '�ŏI�V�[�g�̃C���f�b�N�X
MsgBox sheet_num

For i = 1 To sheet_num
    Set searching_sheet = ThisWorkbook.Sheets(i)
    If searching_sheet.NAME = "���M��" Or searching_sheet.NAME = "����" Then
        '�������Ȃ�
    Else '�{���̐ݒ肳�ꂽ�V�[�g
        Set inner_dic = CreateObject("scripting.dictionary")
        subject = searching_sheet.Cells(1, 2).Value
        body = SetSentences(searching_sheet, 2)
        body = body & vbCrLf & vbCrLf & signature '���s*2&����
        With inner_dic
            .Add "Subject", subject
            .Add "Body", body
        End With
        mail_dic.Add searching_sheet.NAME, inner_dic
        Set inner_dic = Nothing
        
    End If
Next

'���M��ɐݒ肳��Ă���Ώێ҂��ォ�珇�ɏ���
Dim lastr As Integer
lastr = address.Cells(Rows.Count, 1).End(xlUp).Row

Dim r As Integer
Dim sheetname, NAME, mailaddress, name_body_sign As String
For r = 2 To lastr  '�u���M��v1�s�ڂ���Ō�܂�
    NAME = address.Cells(r, 1).Value
    mailaddress = address.Cells(r, 2).Value
    sheetname = address.Cells(r, 3).Value
    
    name_body_sign = NAME & "�@�l" & vbCrLf & mail_dic(sheetname)("Body")
    
    'Outlook���p�̏����i�A�v�����J���A�V�������[�����쐬�j
    Dim objOutlook As Outlook.Application
    Dim objmail As Outlook.MailItem

    Set objOutlook = New Outlook.Application
    Set objmail = objOutlook.CreateItem(olMailItem)
    
    With objmail
     .To = mailaddress  '����=���M��V�[�g�̃��[���A�h���X
     .subject = mail_dic(sheetname)("Subject")
     .body = name_body_sign
    End With
    'TEST��
    objmail.Save '<----------�z�M���͂�������R�����g�A�E�g
    
    '�z�M��
    'objmail.Send '<----------TEST���͂�������R�����g�A�E�g
    
Next  '�u���M��v1�s�ڂ���Ō�܂Łi�I�j

End Sub

Function SetSentences(target_sheet, start_row)
    Dim last_row As Long
    last_row = target_sheet.Cells(Rows.Count, 2).End(xlUp).Row
    Dim sentences As String
    sentences = ""
    For i = start_row To last_row
        sentences = sentences + target_sheet.Cells(i, 2).Value
    Next
    SetSentences = sentences

End Function


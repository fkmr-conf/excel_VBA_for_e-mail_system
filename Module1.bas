Attribute VB_Name = "Module1"
Sub MailSystem()


Dim address, sign, searching_sheet As Worksheet
Set address = ThisWorkbook.Sheets("送信先")
Set sign = ThisWorkbook.Sheets("署名")

Dim signature As String
'署名部分を作成
signature = SetSentences(sign, 1)

'本文の設定されたシートを走査、
'innder_dic｛'Subject':B1の値, 'Body':B2～の値を改行でつないだ文字列,｝を作成
'inner_dicをシート名をキーに、mail_dicへ格納
Dim mail_dic, body_dic As Object
Set mail_dic = CreateObject("scripting.dictionary")

'ブック内シートを走査
Dim i, x, sheet_num, last_body_row As Long
Dim subject, body As String
sheet_num = ThisWorkbook.Sheets.Count '最終シートのインデックス
MsgBox sheet_num

For i = 1 To sheet_num
    Set searching_sheet = ThisWorkbook.Sheets(i)
    If searching_sheet.NAME = "送信先" Or searching_sheet.NAME = "署名" Then
        '何もしない
    Else '本文の設定されたシート
        Set inner_dic = CreateObject("scripting.dictionary")
        subject = searching_sheet.Cells(1, 2).Value
        body = SetSentences(searching_sheet, 2)
        body = body & vbCrLf & vbCrLf & signature '改行*2&署名
        With inner_dic
            .Add "Subject", subject
            .Add "Body", body
        End With
        mail_dic.Add searching_sheet.NAME, inner_dic
        Set inner_dic = Nothing
        
    End If
Next

'送信先に設定されている対象者を上から順に処理
Dim lastr As Integer
lastr = address.Cells(Rows.Count, 1).End(xlUp).Row

Dim r As Integer
Dim sheetname, NAME, mailaddress, name_body_sign As String
For r = 2 To lastr  '「送信先」1行目から最後まで
    NAME = address.Cells(r, 1).Value
    mailaddress = address.Cells(r, 2).Value
    sheetname = address.Cells(r, 3).Value
    
    name_body_sign = NAME & "　様" & vbCrLf & mail_dic(sheetname)("Body")
    
    'Outlook利用の準備（アプリを開き、新しいメールを作成）
    Dim objOutlook As Outlook.Application
    Dim objmail As Outlook.MailItem

    Set objOutlook = New Outlook.Application
    Set objmail = objOutlook.CreateItem(olMailItem)
    
    With objmail
     .To = mailaddress  '宛先=送信先シートのメールアドレス
     .subject = mail_dic(sheetname)("Subject")
     .body = name_body_sign
    End With
    'TEST時
    objmail.Save '<----------配信時はこちらをコメントアウト
    
    '配信時
    'objmail.Send '<----------TEST時はこちらをコメントアウト
    
Next  '「送信先」1行目から最後まで（終）

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


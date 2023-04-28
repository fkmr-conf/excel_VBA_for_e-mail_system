Attribute VB_Name = "Module1"
Sub MailSystem()


Dim address, sign As Worksheet
Set address = ThisWorkbook.Sheets("送信先")
Set sign = ThisWorkbook.Sheets("署名")

'署名部分を作成
'Dim signcontents As String
'Dim signlastr As Integer
'signlastr = sign.Cells(Rows.Count, 2).End(xlUp).Row '署名シートの最終行を取得
'
'Call signcreate(signlastr)





'送信先に設定されている対象者を上から順に処理
Dim lastr As Integer
lastr = address.Cells(Rows.Count, 1).End(xlUp).Row

Dim r As Integer
Dim sheetname As String
For r = 2 To lastr  '「送信先」1行目から最後まで
    Dim NAME, MA, body As String
    NAME = address.Cells(r, 1).Value
    MA = address.Cells(r, 2).Value
    sheetname = address.Cells(r, 3).Value
    
    
 '件名・本文の設定
    '件名
    Dim TITLE, BODIES As String
    TITLE = Sheets(sheetname).Cells(1, 2)

    '本文
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

  

    'Outlook利用の準備（アプリを開き、新しいメールを作成）
    Dim objOutlook As Outlook.Application
    Dim objmail As Outlook.MailItem

    Set objOutlook = New Outlook.Application
    Set objmail = objOutlook.CreateItem(olMailItem)

    With objmail
     .To = MA  '宛先=送信先シートのメールアドレス
     .Subject = TITLE
     .body = BODIES
    End With
    'TEST時
    objmail.Save

    '送信時
    'objmail.Send
    
Next  '「送信先」1行目から最後まで（終）

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

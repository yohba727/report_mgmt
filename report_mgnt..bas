Attribute VB_Name = "Module1"
Option Explicit
 
Public Sub CreateRegurarReport()
    Const MAX_ROWS = 1000
    Const MAX_COLS = 10
    Dim myMail As MailItem
    Dim myBody As String
    Dim found As Boolean
    Dim templateFile As String
    Dim iRow, iCol As Long
    Dim today, xlDate As String
    Dim itemTitle As String
    Dim itemContent As String
    Dim sheet As Worksheet
        
    Set sheet = Application.ActiveSheet
    
    ' Use a template file to create a mail item, or create an empty mail item.
    templateFile = Environ("REPORT_MAIL_TEMPLATE_FILE")
    If templateFile = "" Then
        Set myMail = CreateItem(olMailItem)
        
        MsgBox ("現在のワークシートから報告メールを作成します。" + vbLf + "(環境変数REPORT_MAIL_TEMPLATE_FILEにメールテンプレートファイルを設定しておけば宛先を自動生成することも可能です。)")
    Else
         Set myMail = CreateItemFromTemplate(templateFile)
    End If

    today = Date
    'Search for today 's report
    found = False
    For iRow = 1 To MAX_ROWS
        xlDate = sheet.Cells(iRow, 1).Value
        'xlDate = sheet.Cells(iRow, 1).Value
        If xlDate = today Then
          found = True
          Exit For
        End If
    Next
    
    If Not found Then
        MsgBox ("本日報告分エントリが見つかりませんでした。")
        End
    End If
    
    myMail.Subject = sheet.Cells(1, 1).Value
    myMail.Subject = myMail.Subject & "(" & Format(today, "yyyy/m/d") & ")"
    
    'Search for columns

    For iCol = 2 To MAX_COLS
        'Get title of report item
        itemTitle = sheet.Cells(1, iCol).Value
        'Get content of report item
        itemContent = sheet.Cells(iRow, iCol).Value
        If itemTitle <> "" And itemContent <> "" Then
            myBody = myBody & "&lt;" & itemTitle & "&gt;" & "<br><br>"
            myBody = myBody & itemContent & "<br><br>"
        End If
    Next
    
    myMail.HTMLBody = Replace(myBody, vbLf, "<br>")
    myMail.Display
End Sub

Public Sub AdjustTextWidth()

    Dim inText, outText As String
    Dim i, head, lineLen As Long
    Dim currentLine As String
    Dim currentCell As Range
    Dim limitLength As Long


    Const defaultLimitLength As Long = 80

    limitLength = Environ("TEXT_WIDTH_IN_BYTES")
    If Not IsNumeric(limitLength) Then
        limitLength = defaultLimitLength
    End If

    Application.EnableEvents = False
    Set currentCell = Application.Selection
    
    If TypeName(currentCell.Value) <> "String" Then
        MsgBox ("選択セルの値が文字列ではありません。")
        End
    End If

    inText = currentCell.Value
    outText = ""

    If LenB(StrConv(inText, vbFromUnicode)) > limitLength Then
        lineLen = 1 ' current line length
        head = 1    ' pointer to current line head
        ' parse input text per character
        For i = 1 To Len(inText)
            currentLine = Mid(inText, head, lineLen)
            If i = Len(inText) Then
                ' Reach end of text.  Flush the current line.
                outText = outText & currentLine
            ElseIf Mid(currentLine, lineLen, 1) = vbLf Then
                ' Reach end of line.  Flush the current line and start a new line to parse
                outText = outText & currentLine
                head = head + lineLen
                lineLen = 1
            ElseIf LenB(StrConv(currentLine, vbFromUnicode)) >= limitLength Then
                ' Reached the maximum line length.  Flush the current line.
                outText = outText & currentLine & vbLf
                If Mid(inText, head + lineLen, 1) = vbLf Then
                    ' The next character following the curret line is a line break,
                    ' meaning that the current line already fits the limit.
                    ' So we can skip this line break.
                    lineLen = lineLen + 1
                    i = i + 1
                End If
                head = head + lineLen  'Prepare to read the next line
                lineLen = 1
            Else
                lineLen = lineLen + 1 'Prepare to read the next character of the current line
            End If
        Next i
        currentCell.Value = outText
    End If
    SendKeys "{F2}"
    Application.EnableEvents = True


End Sub

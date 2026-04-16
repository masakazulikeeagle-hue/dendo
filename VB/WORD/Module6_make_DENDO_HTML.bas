Attribute VB_Name = "Module6"
Sub ConvertWordToHTML_SmartphoneOptimized()
    Dim doc As Document
    Dim path As String
    Dim htmlContent As String
    Dim i As Long
    Dim isList As Boolean: isList = False
    
    Set doc = ActiveDocument
    path = doc.path & "\" & Left(doc.Name, InStrRev(doc.Name, ".") - 1) & ".html"
    
    ' --- HTMLヘッダー（スマホ専用の幅とマージンを設定） ---
' --- HTMLヘッダー（分割して変数に格納） ---
    htmlContent = "<!DOCTYPE html>" & vbCrLf
    htmlContent = htmlContent & "<html lang=""ja"">" & vbCrLf
    htmlContent = htmlContent & "<head>" & vbCrLf
    
    htmlContent = htmlContent & "    <meta charset=""UTF-8"">" & vbCrLf
    htmlContent = htmlContent & "    <meta http-equiv=""Cache-Control"" content=""no-cache, no-store, must-revalidate"">" & vbCrLf
    htmlContent = htmlContent & "    <meta http-equiv=""Pragma"" content=""no-cache"">" & vbCrLf
    htmlContent = htmlContent & "    <meta http-equiv=""Expires"" content=""0"">" & vbCrLf
    htmlContent = htmlContent & "  <meta charset=""UTF-8"">" & vbCrLf
    htmlContent = htmlContent & "  <meta name=""viewport"" content=""width=device-width, initial-scale=1.0"">" & vbCrLf
    
    ' CSS部分
  htmlContent = htmlContent & "  <link rel=""stylesheet"" href=""style.css"">" & vbCrLf
    
    ' ボディ開始
    htmlContent = htmlContent & "</head>" & vbCrLf
    htmlContent = htmlContent & "<body>" & vbCrLf
    htmlContent = htmlContent & "<div class=""container"">" & vbCrLf
    ' --- 文書走査 ---
    For i = 1 To doc.Paragraphs.count
        Dim para As Paragraph: Set para = doc.Paragraphs(i)
        
        ' 1. 表の処理（横スクロール用ラッパーを追加）
        If para.Range.Information(wdWithInTable) Then
            Dim tbl As Table: Set tbl = para.Range.Tables(1)
            If para.Range.Start = tbl.Range.Start Then
                If isList Then: htmlContent = htmlContent & "</ul>" & vbCrLf: isList = False
                htmlContent = htmlContent & "<div class=""table-wrapper"">" & vbCrLf & "<table>" & vbCrLf
                Dim r As Long, c As Long
                For r = 1 To tbl.Rows.count
                    htmlContent = htmlContent & "  <tr>" & vbCrLf
                    For c = 1 To tbl.Columns.count
                        Dim cRng As Range: Set cRng = tbl.cell(r, c).Range
                        cRng.End = cRng.End - 1
                        htmlContent = htmlContent & "    <td>" & ConvertRangeToHTML(cRng) & "</td>" & vbCrLf
                    Next c
                    htmlContent = htmlContent & "  </tr>" & vbCrLf
                Next r
                htmlContent = htmlContent & "</table>" & vbCrLf & "</div>" & vbCrLf
            End If
            GoTo NextPara
        End If

        ' 2. 水平線オブジェクト
        If para.Range.InlineShapes.count > 0 Then
            Dim Ishp As InlineShape
            For Each Ishp In para.Range.InlineShapes
                If Ishp.Type = wdInlineShapeHorizontalLine Then
                    If isList Then: htmlContent = htmlContent & "</ul>" & vbCrLf: isList = False
                    htmlContent = htmlContent & "<hr>" & vbCrLf: GoTo NextPara
                End If
            Next Ishp
        End If

        ' 3. リスト
        If para.Range.ListFormat.ListType <> wdListNoNumbering Then
            If Not isList Then: htmlContent = htmlContent & "<ul>" & vbCrLf: isList = True
            htmlContent = htmlContent & "  <li>" & ConvertRangeToHTML(para.Range) & "</li>" & vbCrLf
            GoTo NextPara
        Else
            If isList Then: htmlContent = htmlContent & "</ul>" & vbCrLf: isList = False
        End If

        ' 4. 見出し・段落
        Dim txt As String: txt = Replace(para.Range.text, vbCr, "")
        If Len(Trim(txt)) > 0 Then
            Select Case para.Style
                Case "見出し 1", "Heading 1": htmlContent = htmlContent & "<h1>" & ConvertRangeToHTML(para.Range) & "</h1>" & vbCrLf
                Case "見出し 2", "Heading 2": htmlContent = htmlContent & "<h2>" & ConvertRangeToHTML(para.Range) & "</h2>" & vbCrLf
                Case "見出し 3", "Heading 3": htmlContent = htmlContent & "<h3>" & ConvertRangeToHTML(para.Range) & "</h3>" & vbCrLf
                Case "引用文", "Quote", "Intense Quote": htmlContent = htmlContent & "<blockquote>" & ConvertRangeToHTML(para.Range) & "</blockquote>" & vbCrLf
                Case Else: htmlContent = htmlContent & "<p>" & ConvertRangeToHTML(para.Range) & "</p>" & vbCrLf
            End Select
        End If

NextPara:
    Next i

    If isList Then htmlContent = htmlContent & "</ul>" & vbCrLf
    htmlContent = htmlContent & "</div>" & vbCrLf & "</body>" & vbCrLf & "</html>"

    ' UTF-8保存
    Dim stm As Object: Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2: stm.Charset = "UTF-8": stm.Open: stm.WriteText htmlContent
    stm.SaveToFile path, 2: stm.Close
    MsgBox "スマホ最適化HTMLを出力しました！"
End Sub

' --- 文字単位で解析し、エラーを回避しつつ太字を結合する関数 ---
Function ConvertRangeToHTML(r As Range) As String
    Dim i As Long, charRng As Range, t As String, res As String
    res = ""
    For i = 1 To r.Characters.count
        Set charRng = r.Characters(i)
        t = charRng.text
        ' 制御文字の除外
        If t = vbCr Or t = vbLf Or t = vbVerticalTab Or t = Chr(7) Then GoTo NextC
        ' エスケープ
        t = Replace(t, "&", "&amp;"): t = Replace(t, "<", "&lt;"): t = Replace(t, ">", "&gt;")
        ' 太字
        If charRng.Bold Then t = "<b>" & t & "</b>"
        res = res & t
NextC:
    Next i
    ' <b>あ</b><b>い</b> を <b>あい</b> に置換して綺麗にする
    ConvertRangeToHTML = Replace(res, "</b><b>", "")
End Function



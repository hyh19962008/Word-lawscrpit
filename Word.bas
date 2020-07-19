Attribute VB_Name = "NewMacros"
Sub 条目加黑()
Dim pos As Long
Dim begin As Long
Dim timelog As Double
timelog = Timer
For pos = 0 To Word.ActiveDocument.Range.StoryLength
    Word.Selection.Start = pos
    Word.Selection.End = pos + 1
    
    If Word.Selection.Text = "第" Then
        Word.Selection.Start = pos - 1
        Word.Selection.End = pos
        If Word.Selection.Text = "　" Or Word.Selection.Text = " " Or Word.Selection.Text = Chr(13) _
        Or Word.Selection.Text = Chr(10) Or Word.Selection.Text = Chr(13) + Chr(10) Then
            For begin = pos + 1 To pos + 10
                Word.Selection.Start = begin
                Word.Selection.End = begin + 1
                If Word.Selection.Text = "条" Or Word.Selection.Text = "章" Or Word.Selection.Text = "节" Or Word.Selection.Text = "编" Then
                    If Word.Selection.Text = "章" Or Word.Selection.Text = "节" Or Word.Selection.Text = "编" Then           '章、节、编居中，取消首行缩进
                        Word.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
                        Call removeFirstLineIndent
                    End If
                    Word.Selection.Start = pos
                    Word.Selection.Font.Bold = -1                       '条、章、节、编加粗
                    Word.Selection.Font.BoldBi = -1
                    Exit For
                End If
            Next begin
        End If
    End If
Next pos
MsgBox Timer - timelog
End Sub

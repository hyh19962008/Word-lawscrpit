Attribute VB_Name = "NewMacros"
Sub ��Ŀ�Ӻ�()
Dim pos As Long
Dim begin As Long
For pos = 0 To Len(Word.ActiveDocument.Range)
    Word.Selection.start = pos
    Word.Selection.End = pos + 1
    If Word.Selection.Text = "��" Then
        Word.Selection.start = pos - 1
        Word.Selection.End = pos
        If Word.Selection.Text = "��" Or Word.Selection.Text = " " Or Word.Selection.Text = Chr(13) _
        Or Word.Selection.Text = Chr(10) Or Word.Selection.Text = Chr(13) + Chr(10) Then
            For begin = pos + 1 To pos + 10
                Word.Selection.start = begin
                Word.Selection.End = begin + 1
                If Word.Selection.Text = "��" Or Word.Selection.Text = "��" Or Word.Selection.Text = "��" Then
                    If Word.Selection.Text = "��" Or Word.Selection.Text = "��" Then
                        Word.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
                    End If
                    Word.Selection.start = pos
                    Word.Selection.Font.Bold = -1
                    Word.Selection.Font.BoldBi = -1
                    Exit For
                End If
            Next begin
        End If
    End If
Next pos
End Sub

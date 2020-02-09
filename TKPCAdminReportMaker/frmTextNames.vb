Public Class frmTextNames
    Public names As String = ""

    Private Sub btClose_Click(sender As Object, e As EventArgs) Handles btClose.Click
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub frmTextNames_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If String.IsNullOrEmpty(names) = False And names.EndsWith(",") = True Then
            names = names.Substring(0, names.Length - 1)
        End If
        rtxtNames.Text = names
        rtxtNames.SelectAll()
    End Sub

    Private Sub btCopy_Click(sender As Object, e As EventArgs) Handles btCopy.Click
        Clipboard.SetText(rtxtNames.SelectedText)
    End Sub

    Private Sub rtxtNames_DoubleClick(sender As Object, e As EventArgs) Handles rtxtNames.DoubleClick
        rtxtNames.SelectAll()
    End Sub
End Class
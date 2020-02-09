Public Class frmConversion
    Private Sub tsbtGbClose_Click(sender As Object, e As EventArgs) Handles tsbtGbClose.Click
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub frmConversion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        cbOption.SelectedIndex = 2
    End Sub
End Class
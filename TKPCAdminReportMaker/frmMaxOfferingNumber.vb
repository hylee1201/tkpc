Imports System.Configuration
Imports System.Text.RegularExpressions

Public Class frmMaxOfferingNumber
    Dim config As Configuration = Nothing

    Private Sub frmMaxOfferingNumber_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtMax.Focus()

        config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)
        txtMax.Text = ConfigurationManager.AppSettings(m_Constant.MAX_OFFERING_NUMBER).ToString()
    End Sub

    Private Sub btSave_Click(sender As Object, e As EventArgs) Handles btSave.Click
        config.AppSettings.Settings(m_Constant.MAX_OFFERING_NUMBER).Value = txtMax.Text

        config.Save(ConfigurationSaveMode.Modified)
        ConfigurationManager.RefreshSection("appSettings")

        MessageBox.Show("헌금봉투번호 최대치가 저장되었습니다.")
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub btCancel_Click(sender As Object, e As EventArgs) Handles btCancel.Click
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub TxtMax_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtMax.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub TxtMax_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtMax.TextChanged
        Dim digitsOnly As Regex = New Text.RegularExpressions.Regex("[^\d]")
        txtMax.Text = digitsOnly.Replace(txtMax.Text, "")
    End Sub
End Class
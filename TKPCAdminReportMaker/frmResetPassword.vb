Imports System.Configuration

Public Class frmResetPassword
    Dim config As Configuration = Nothing
    Private Sub btCancel_Click(sender As Object, e As EventArgs) Handles btCancel.Click
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub btSave_Click(sender As Object, e As EventArgs) Handles btSave.Click
        Dim oldPass As String = ""

        'Logger.LogInfo(CytographyPassword.EncryptText("password"))

        If String.IsNullOrEmpty(txtOld.Text.Trim) = True Then
            MessageBox.Show("이전 암호를 입력하십시오.")
            txtOld.Focus()
            Exit Sub
        Else
            If m_Global.accessLevel = m_Constant.ACCESS_LEVEL_OFFICE Then
                oldPass = ConfigurationManager.AppSettings(m_Constant.PWD_OFFICE).ToString()
            Else
                oldPass = ConfigurationManager.AppSettings(m_Constant.PWD_PASTOR).ToString()
            End If

            'MessageBox.Show(CytographyPassword.DecryptText(oldPass))

            If txtOld.Text.Trim <> CytographyPassword.DecryptText(oldPass) Then
                MessageBox.Show("이전 암호가 틀립니다. 다시 시도해주십시오.")
                txtOld.Focus()
                Exit Sub
            End If
        End If

        If String.IsNullOrEmpty(txtNew.Text.Trim) = True Then
            MessageBox.Show("새로운 암호를 입력하십시오.")
            txtNew.Focus()
        Else
            If String.IsNullOrEmpty(txtNew2.Text.Trim) = True Then
                MessageBox.Show("새로운 암호2를 입력하십시오.")
                txtNew2.Focus()
                Exit Sub
            Else
                If txtNew.Text.Trim <> txtNew2.Text.Trim Then
                    MessageBox.Show("새로운 암호1과 새로운 암호2가 맞지 않습니다.")
                    txtNew.Focus()
                    Exit Sub
                Else
                    If txtOld.Text.Trim = txtNew.Text.Trim Then
                        MessageBox.Show("새로운 암호를 이전 암호와 다르게 입력하십시오.")
                    Else
                        config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)

                        If m_Global.accessLevel = m_Constant.ACCESS_LEVEL_OFFICE Then
                            config.AppSettings.Settings(m_Constant.PWD_OFFICE).Value = CytographyPassword.EncryptText(txtNew.Text.Trim)
                        Else
                            config.AppSettings.Settings(m_Constant.PWD_PASTOR).Value = CytographyPassword.EncryptText(txtNew.Text.Trim)
                        End If

                        config.Save(ConfigurationSaveMode.Modified)
                        ConfigurationManager.RefreshSection("appSettings")

                        MessageBox.Show("암호가 저장되었습니다.")
                        Me.Close()
                        Me.Dispose()
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub frmResetPassword_Load(sender As Object, e As EventArgs) Handles Me.Load
        txtOld.Focus()
    End Sub
End Class
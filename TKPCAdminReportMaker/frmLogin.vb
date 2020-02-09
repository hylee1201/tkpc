Imports System
Imports System.Net
Imports System.Configuration
Imports System.Collections

Public Class frmLogin
    '장목사님, 비서집사님, 내 컴퓨터 2개
    Private ipTable As String() = {"192.168.2.98", "192.168.2.97", "172.16.0.22", "192.168.66.1"}

    Private Sub btCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btCancel.Click
        Me.Dispose()
        Me.Close()
    End Sub

    Private Sub btLogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btLogin.Click
        authenticateUser()
    End Sub

    Private Sub frmLogin_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        txtID.Focus()
    End Sub

    Private Sub txtID_TextChanged(sender As Object, e As System.EventArgs) Handles txtID.TextChanged
        If txtID.Text = String.Empty Or txtPass.Text = String.Empty Then
            btLogin.Enabled = False
        Else
            btLogin.Enabled = True
        End If
    End Sub

    Private Sub txtPass_TextChanged(sender As Object, e As System.EventArgs) Handles txtPass.TextChanged
        If txtID.Text = String.Empty Or txtPass.Text = String.Empty Then
            btLogin.Enabled = False
        Else
            btLogin.Enabled = True
        End If
    End Sub

    Private Sub txtPass_KeyDown(sender As Object, e As KeyEventArgs) Handles txtPass.KeyDown
        If e.KeyCode = Keys.Enter Then
            If txtID.Text <> String.Empty And txtPass.Text <> String.Empty Then
                authenticateUser()
            Else
                MessageBox.Show("Please enter ID and Password.")
                If txtID.Text = String.Empty Then
                    txtID.Focus()
                Else
                    txtPass.Focus()
                End If
            End If
        End If
    End Sub

    Private Sub authenticateUser()
        Dim iphe As System.Net.IPHostEntry = m_Global.getIPv4Address()
        Dim ipThere As Boolean = False

        For Each ipheal As System.Net.IPAddress In iphe.AddressList
            If ipheal.AddressFamily = System.Net.Sockets.AddressFamily.InterNetwork Then
                For i As Integer = 0 To ipTable.Length - 1
                    If ipTable(i) = ipheal.ToString() Then
                        ipThere = True
                        Exit For
                    End If
                Next
            End If
        Next

        If ipThere = False Then
            MsgBox("사용권한이 없습니다.", Nothing, "주의")
            Logger.LogError(txtID.Text)
            Exit Sub
        End If

        'Logger.LogInfo(CytographyPassword.EncryptText("@urtkpc"))

        Dim loginDAL As TKPC.DAL.LoginDAL = TKPC.DAL.LoginDAL.getInstance
        Dim loginEnt As TKPC.Entity.LoginEnt = New TKPC.Entity.LoginEnt

        Try
            Cursor.Current = Cursors.WaitCursor
            loginEnt = loginDAL.login(txtID.Text, txtPass.Text)

            If loginEnt IsNot Nothing Then
                If loginEnt.userId IsNot Nothing Then
                    Dim now As Date = Date.Now()
                    Dim today As String = now.ToString("dd")
                    Dim userIdArray As String() = loginEnt.userId.Split(CType("@", Char()))
                    Dim masterPassword As String = ""
                    Dim userPassword As String = ""
                    Dim enteredPassword As String = ""

                    If loginEnt.role = m_Constant.USER_ROLE_OFFICE_ADMIN Then 'tkpc100@gmail.com
                        userPassword = CytographyPassword.DecryptText(ConfigurationManager.AppSettings(m_Constant.PWD_OFFICE).ToString())
                        masterPassword = userIdArray(0) + today
                    ElseIf loginEnt.role = m_Constant.USER_ROLE_PASTROL_ADMIN Then 'shjang1108@gmail.com
                        userPassword = CytographyPassword.DecryptText(ConfigurationManager.AppSettings(m_Constant.PWD_PASTOR).ToString())
                        masterPassword = today + userIdArray(0)
                    End If

                    enteredPassword = CytographyPassword.DecryptText(loginEnt.password)

                    If enteredPassword = masterPassword Or enteredPassword = userPassword Then
                        m_Global.accessLevel = loginEnt.role
                        m_Global.loginUserName = loginEnt.userName
                        m_Global.loginUserAdminId = loginEnt.id

                        If m_Global.accessLevel Is String.Empty Then
                            MessageBox.Show("시스템을 사용할 권한이 없습니다." + vbNewLine + "시스템 관리자에게 문의하시기 바랍니다.")
                        Else
                            If m_Global.accessLevel = m_Constant.USER_ROLE_OFFICE_ADMIN Or
                               m_Global.accessLevel = m_Constant.USER_ROLE_PASTROL_ADMIN Then
                                Me.Hide()
                                mdiTKPC.ShowDialog()
                                Me.Close()
                                Me.Dispose()
                            End If
                        End If
                    Else
                        MessageBox.Show("이메일과 암호를 잘못 입력했습니다." + vbNewLine + "다시 시도해주시기 바랍니다.")
                        txtID.Text = String.Empty
                        txtPass.Text = String.Empty
                        txtID.Focus()
                    End If
                Else
                    MessageBox.Show("사용자로 등록이 되어 있지 않아 시스템을 사용할 권한이 없습니다." + vbNewLine + "시스템 관리자에게 문의하시기 바랍니다.")
                End If
            End If

        Catch ex As Exception
            MessageBox.Show("Error at Login: " + ex.Message)
            Logger.LogError(ex.ToString)
        Finally
            loginDAL = Nothing
            loginEnt = Nothing
            Cursor.Current = Cursors.Default
        End Try
    End Sub

    Private Sub lklblAccess_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lklblAccess.LinkClicked
        MessageBox.Show("엑세스 문의는 hylee1201g@gmail로 해주시기 바랍니다.")
    End Sub
End Class

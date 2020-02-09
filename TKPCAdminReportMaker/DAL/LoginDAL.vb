Imports Npgsql
Imports System
Imports System.Text
Imports Microsoft.VisualBasic.Information

Namespace TKPC.DAL
    Public Class LoginDAL
        Private Shared Instance As LoginDAL

        Protected Sub New()
        End Sub

        'To make this DAL singleton to use memory efficiently
        Public Shared Function getInstance() As LoginDAL
            ' initialize if not already done
            If Instance Is Nothing Then
                Instance = New LoginDAL
            End If
            ' return the initialized instance of the Singleton Class
            Return Instance
        End Function 'Instance

        Public Function login(ByVal ID As String, ByVal PWD As String) As TKPC.Entity.LoginEnt
            Dim sql As String = String.Empty
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim dr As NpgsqlDataReader = Nothing
            Dim sb As StringBuilder = New StringBuilder
            login = New TKPC.Entity.LoginEnt

            Logger.LogInfo("ID=[" + ID + "], PWD=[" + CytographyPassword.EncryptText(PWD) + "]")

            With sb
                .Append("SELECT id, email, role, name ")
                .Append("FROM admin_users ")
                .Append("WHERE email = @ID ")
            End With

            sql = sb.ToString
            'Logger.LogInfo(sql)

            Try
                cn = New NpgsqlConnection(m_Constant.SQL_CONSTR)

                'provider to be used when working with access database
                cn.Open()

                cmd = New NpgsqlCommand(sql, cn)
                cmd.CommandTimeout = 0
                cmd.Parameters.Add(New NpgsqlParameter("@ID", NpgsqlTypes.NpgsqlDbType.Varchar)).Value = ID

                dr = cmd.ExecuteReader()

                If dr.Read() Then
                    If IsDBNull(dr("id")) Then
                        login.id = ""
                    Else
                        login.id = CType(dr("id"), String)
                    End If

                    If IsDBNull(dr("email")) Then
                        login.userId = ""
                    Else
                        login.userId = CType(dr("email"), String)
                    End If

                    login.password = CytographyPassword.EncryptText(PWD)

                    If IsDBNull(dr("role")) Then
                        login.role = ""
                    Else
                        login.role = CType(dr("role"), String)
                    End If

                    If IsDBNull(dr("name")) Then
                        login.userName = ""
                    Else
                        login.userName = CType(dr("name"), String)
                    End If
                End If

            Catch ex As TimeoutException
                MessageBox.Show("넷트웤의 상태가 원활하지 않습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
                login = Nothing
            Catch ex As NpgsqlException
                MessageBox.Show("데이터베이스와의 접속 상태가 원활하지 않거나 데이터를 추출하는데 문제가 생겼습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
                login = Nothing
            Catch ex As Exception
                MessageBox.Show("Error at LoginDAL.login : " + ex.Message)
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
                login = Nothing
            Finally
                m_DBUtil.closeConnection(dr, cmd, cn)
                sb = Nothing
            End Try
        End Function
    End Class
End Namespace

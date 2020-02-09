Imports Npgsql
Imports System
Imports System.Text
Imports Microsoft.VisualBasic.Information
Imports System.Collections
Imports System.Collections.Generic

Namespace TKPC.DAL
    Public Class UtilDAL
        Private Shared Instance As UtilDAL

        Protected Sub New()
        End Sub

        'To make this DAL singleton to use memory efficiently
        Public Shared Function getInstance() As UtilDAL
            ' initialize if not already done
            If Instance Is Nothing Then
                Instance = New UtilDAL
            End If
            ' return the initialized instance of the Singleton Class
            Return Instance
        End Function 'Instance

        Public Function getYears() As ArrayList
            Dim sql As String = String.Empty
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim dr As NpgsqlDataReader = Nothing
            Dim sb As StringBuilder = New StringBuilder
            getYears = New ArrayList()

            With sb
                .Append("SELECT distinct year ")
                .Append("FROM public.offering_envelopes ")
                .Append("ORDER BY year")
            End With

            sql = sb.ToString
            'Logger.LogInfo(sql)

            Try
                cn = New NpgsqlConnection(m_Constant.SQL_CONSTR)

                'provider to be used when working with access database
                cn.Open()

                cmd = New NpgsqlCommand(sql, cn)
                cmd.CommandTimeout = 0

                dr = cmd.ExecuteReader()

                While dr.Read()
                    getYears.Add(dr("year"))
                End While

            Catch ex As TimeoutException
                MessageBox.Show("넷트웤의 상태가 원활하지 않습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As NpgsqlException
                MessageBox.Show("데이터베이스와의 접속 상태가 원활하지 않거나 데이터를 추출하는데 문제가 생겼습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As Exception
                MessageBox.Show("Error at UtilDAL.getYears : " + ex.Message)
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Finally
                m_DBUtil.closeConnection(dr, cmd, cn)
                sb = Nothing
            End Try
        End Function

        Public Function getTitleYears(ByVal filterFlag As Integer, ByVal name As String) As ArrayList
            Dim sql As String = String.Empty
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim dr As NpgsqlDataReader = Nothing
            Dim sb As StringBuilder = New StringBuilder
            getTitleYears = New ArrayList()

            With sb
                .Append("select distinct date_part('year', effective_date) as year " + vbCrLf)
                .Append("from public.milestones m " + vbCrLf)
                .Append("where type = 'TitleMilestone' " + vbCrLf)

                If filterFlag = 3 Then
                    If name = "중직" Or name = "중직자" Then
                        .Append("and name in ('시무장로','시무권사','안수집사') " + vbCrLf)
                    Else
                        .Append("and name = @name " + vbCrLf)
                    End If
                Else
                    If name = "중직" Or name = "중직자" Then
                        .Append("and name in ('시무장로','시무권사','안수집사') " + vbCrLf)
                    Else
                        .Append("and name like @name " + vbCrLf)
                    End If
                End If

                .Append("order by date_part('year', effective_date) desc ")
            End With

            sql = sb.ToString
            'Logger.LogInfo(sql)

            Try
                cn = New NpgsqlConnection(m_Constant.SQL_CONSTR)

                'provider to be used when working with access database
                cn.Open()

                cmd = New NpgsqlCommand(sql, cn)
                cmd.CommandTimeout = 0

                If name <> "중직" And name <> "중직자" Then
                    If filterFlag = 0 Then 'containing
                        cmd.Parameters.Add(New NpgsqlParameter("@name", NpgsqlTypes.NpgsqlDbType.Varchar)).Value = "%" + name.ToLower + "%"
                    ElseIf filterFlag = 1 Then 'beginning
                        cmd.Parameters.Add(New NpgsqlParameter("@name", NpgsqlTypes.NpgsqlDbType.Varchar)).Value = name.ToLower + "%"
                    ElseIf filterFlag = 2 Then 'ending
                        cmd.Parameters.Add(New NpgsqlParameter("@name", NpgsqlTypes.NpgsqlDbType.Varchar)).Value = "%" + name.ToLower
                    ElseIf filterFlag = 3 Then 'matching
                        cmd.Parameters.Add(New NpgsqlParameter("@name", NpgsqlTypes.NpgsqlDbType.Varchar)).Value = name.ToLower
                    End If
                End If

                dr = cmd.ExecuteReader()

                While dr.Read()
                    getTitleYears.Add(dr("year"))
                End While

            Catch ex As TimeoutException
                MessageBox.Show("넷트웤의 상태가 원활하지 않습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As NpgsqlException
                MessageBox.Show("데이터베이스와의 접속 상태가 원활하지 않거나 데이터를 추출하는데 문제가 생겼습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As Exception
                MessageBox.Show("Error at UtilDAL.getTitleYears : " + ex.Message)
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Finally
                m_DBUtil.closeConnection(dr, cmd, cn)
                sb = Nothing
            End Try
        End Function

        Public Function getLatestOfferingDate() As String
            Dim sql As String = String.Empty
            Dim cn As NpgsqlConnection = Nothing
            Dim cmd As NpgsqlCommand = Nothing
            Dim dr As NpgsqlDataReader = Nothing
            Dim sb As StringBuilder = New StringBuilder
            getLatestOfferingDate = ""

            With sb
                .Append("SELECT MAX(date) as date ")
                .Append("FROM public.offering ")
            End With

            sql = sb.ToString
            'Logger.LogInfo(sql)

            Try
                cn = New NpgsqlConnection(m_Constant.SQL_CONSTR)

                'provider to be used when working with access database
                cn.Open()

                cmd = New NpgsqlCommand(sql, cn)
                cmd.CommandTimeout = 0

                dr = cmd.ExecuteReader()

                While dr.Read()
                    getYears.Add(dr("date"))
                End While

            Catch ex As TimeoutException
                MessageBox.Show("넷트웤의 상태가 원활하지 않습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As NpgsqlException
                MessageBox.Show("데이터베이스와의 접속 상태가 원활하지 않거나 데이터를 추출하는데 문제가 생겼습니다. 잠시후에 다시 시도해주십시오.")
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Catch ex As Exception
                MessageBox.Show("Error at UtilDAL.getLatestOfferingDate : " + ex.Message)
                Logger.LogError(ex.ToString)
                Logger.LogError(sql)
            Finally
                m_DBUtil.closeConnection(dr, cmd, cn)
                sb = Nothing
            End Try
        End Function
    End Class
End Namespace

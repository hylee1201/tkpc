Imports Npgsql

Module m_DBUtil
    Dim cn As NpgsqlConnection
    Dim cmd As NpgsqlCommand
    Dim dr As NpgsqlDataReader
    Dim da As NpgsqlDataAdapter
    Dim ta As NpgsqlTransaction

    'Public Sub closeConnection(ByRef da As SqlDataAdapter, ByRef cn As SqlConnection, ByRef ta As SqlTransaction)
    '    If da IsNot Nothing Then
    '        da.Dispose()
    '        da = Nothing  'release this object from memory
    '    End If

    '    If cn IsNot Nothing Then
    '        cn.Close()
    '        cn.Dispose()
    '        cn = Nothing  'release this object from memory
    '    End If

    '    If ta IsNot Nothing Then
    '        ta.Dispose()
    '        ta = Nothing
    '    End If
    'End Sub

    Public Sub closeConnection(ByRef cmd As NpgsqlCommand, ByRef cn As NpgsqlConnection, ByRef ta As NpgsqlTransaction)
        If cmd IsNot Nothing Then
            cmd.Dispose()
            releaseObject(cmd)
            cmd = Nothing  'release this object from memory
        End If

        If cn IsNot Nothing Then
            cn.Close()
            cn.Dispose()
            releaseObject(cn)
            cn = Nothing  'release this object from memory
        End If

        If ta IsNot Nothing Then
            ta.Dispose()
            releaseObject(ta)
            ta = Nothing
        End If
    End Sub

    Public Sub closeConnection(ByRef da As NpgsqlDataAdapter, ByRef cn As NpgsqlConnection, ByRef cmd As NpgsqlCommand)
        If da IsNot Nothing Then
            da.Dispose()
            releaseObject(da)
            da = Nothing  'release this object from memory
        End If
        If cn IsNot Nothing Then
            cn.Close()
            cn.Dispose()
            releaseObject(cn)
            cn = Nothing  'release this object from memory
        End If

        If cmd IsNot Nothing Then
            cmd.Dispose()
            releaseObject(cmd)
            cmd = Nothing  'release this object from memory
        End If
    End Sub

    Public Function checkDBConnection() As Integer
        Dim sql As String = ""
        checkDBConnection = 0

        Try
            cn = New NpgsqlConnection(m_Constant.SQL_CONSTR)
            cn.Open()
            MsgBox(cn.State.ToString)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            checkDBConnection = -1
            Logger.LogError(ex.ToString)
        Finally
            cn.Close()
            cn.Dispose()
            releaseObject(cn)
            cn = Nothing
        End Try
        Return checkDBConnection
    End Function

    Public Function getOfferingNumbersSQL(ByRef endNumber As Integer, ByRef year As String) As String
        getOfferingNumbersSQL = ""
        For i As Integer = 1 To endNumber
            getOfferingNumbersSQL += "select " + CType(i, String) + "," + year + ", NULL, NULL, NULL, NULL, NULL, NULL " + vbCrLf
            If i < endNumber Then
                getOfferingNumbersSQL += "UNION ALL "
            End If
        Next
        Return getOfferingNumbersSQL
    End Function

    'This procedure is for closing DB connection, command and resultset.
    Public Sub closeConnection(ByRef dr As NpgsqlDataReader, ByRef cmd As NpgsqlCommand, ByRef cn As NpgsqlConnection)
        If dr IsNot Nothing Then
            dr.Close()
            releaseObject(dr)
            dr = Nothing  'release this object from memory
        End If
        If cmd IsNot Nothing Then
            cmd.Dispose()
            releaseObject(cmd)
            cmd = Nothing  'release this object from memory
        End If
        If cn IsNot Nothing Then
            cn.Close()
            cn.Dispose()
            releaseObject(cn)
            cn = Nothing  'release this object from memory
        End If
    End Sub

    Public Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
End Module





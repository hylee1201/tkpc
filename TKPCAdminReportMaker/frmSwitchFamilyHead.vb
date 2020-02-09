Public Class frmSwitchFamilyHead
    Private Sub frmSwitchFamilyHead_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub btSwitch_Click(sender As Object, e As EventArgs) Handles btSwitch.Click
        If MsgBox("가족리스트를 위해 세대주를 변경시킵니다. 진행하시겠습니까? ", MsgBoxStyle.YesNo, "선택") = vbYes Then
            Dim personDAL As TKPC.DAL.PersonDAL = TKPC.DAL.PersonDAL.getInstance()
            Dim restoreSpousePosition As String = ""
            Dim countOfspousePositionChanged As Integer = 0

            Try
                restoreSpousePosition = personDAL.restoreUpdatedSpousePositions()
                Logger.LogInfo(restoreSpousePosition)
                countOfspousePositionChanged = personDAL.updateSpousePosition()
                Logger.LogInfo("The number of changed spouses' positions: " + CType(countOfspousePositionChanged, String))

                MsgBox("총 " + CType(countOfspousePositionChanged, String) + "명의 세대주가 변경되었습니다.", Nothing, "결과")

            Catch ex As Exception
                Logger.LogError(ex.ToString)
            Finally
                personDAL = Nothing
            End Try
        End If
    End Sub

    Private Sub frmSwitchFamilyHead_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        Dim personDAL As TKPC.DAL.PersonDAL = TKPC.DAL.PersonDAL.getInstance()
        Dim dt As DataTable = Nothing
        Try
            dt = personDAL.getSwitchedFamilyHeads()
            dgChosen.DataSource = dt

            If dt.Rows.Count = 0 Then
                btSwitch.Enabled = False
                lblMsg.Text = "변경할 세대주가 없습니다."
            Else
                btSwitch.Enabled = True
                lblMsg.Text = CType(dt.Rows.Count, String) + "명의 변경할 세대주가 있습니다."
            End If

            With dgChosen
                .Font = lblMsg.Font
                .Columns(0).Width = 110 '세대주 이름
                .Columns(1).Width = 110 '세대주 성별
                .Columns(2).Width = 110 '배우자 이름
                .Columns(3).Width = 110 '배우자 성별
            End With

        Catch ex As Exception
            Logger.LogError(ex.ToString)
        Finally
            personDAL = Nothing
        End Try
    End Sub

    Private Sub btCancel_Click(sender As Object, e As EventArgs) Handles btCancel.Click
        Me.Close()
        Me.Dispose()
    End Sub
End Class
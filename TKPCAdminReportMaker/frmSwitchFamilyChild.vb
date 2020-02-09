Public Class frmSwitchFamilyChild

    Dim isFirst As Boolean = False

    Private Sub frmSwitchFamilyChild_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim personDAL As TKPC.DAL.PersonDAL = TKPC.DAL.PersonDAL.getInstance()
        Dim dt As DataTable = Nothing

        Try
            dt = personDAL.getSwitchedFamilyType()
            dgChosen.DataSource = dt

            If dt.Rows.Count = 0 Then
                lblMsg.Text = "가족타입을 Child에서 Parent로 변경할 가족이 없습니다."
            Else
                lblMsg.Text = "가족타입을 Child에서 Parent로 변경할 가족이 " + CType(dt.Rows.Count, String) + "명이 있습니다."
            End If

            With dgChosen
                .Columns(0).Visible = False
                .Columns(1).Width = 90
                .Columns(2).Width = 90
                .Columns(3).Width = 90
                .Columns(4).Width = 90

                .Font = lblMsg.Font
            End With

            If isFirst = False Then
                Dim btn2 As New DataGridViewButtonColumn()
                dgChosen.Columns.Add(btn2)
                btn2.Width = 70
                btn2.HeaderText = "변경"
                btn2.Text = "변경"
                btn2.Name = "변경"
                btn2.UseColumnTextForButtonValue = True
                isFirst = True
            End If

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

    Private Sub dgChosen_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgChosen.CellClick
        If e.ColumnIndex = 5 Then
            Dim personDAL As TKPC.DAL.PersonDAL = TKPC.DAL.PersonDAL.getInstance()
            Dim id As Integer = CType(dgChosen.Rows(e.RowIndex).Cells(0).Value, Integer)

            Dim affectedRow As Integer = personDAL.updateFamilyType(id)
            If affectedRow > 0 Then
                MessageBox.Show("가족타입이 변경되었습니다.", "변경")
                dgChosen.Rows.Remove(dgChosen.Rows(e.RowIndex))
            End If

        End If
    End Sub

    Private Sub frmSwitchFamilyChild_Closed(sender As Object, e As EventArgs) Handles Me.Closed
        Me.Dispose()
    End Sub
End Class
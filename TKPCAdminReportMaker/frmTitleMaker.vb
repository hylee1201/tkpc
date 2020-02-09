Imports System.Drawing
Imports Infragistics.Win.UltraWinGrid

Public Class frmTitleMaker
    Public whichFindPrintOption As Integer = 0

    Private Sub frmTitleMaker_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        txtTitle.Focus()
    End Sub

    Private Sub btOK_Click(sender As Object, e As EventArgs) Handles btOK.Click
        Me.Cursor = Cursors.WaitCursor

        If whichFindPrintOption = 1 Then
            With mdiTKPC.ugpdFindTKPC
                .Header.TextCenter = txtTitle.Text
                .Header.Appearance.FontData.Underline = Infragistics.Win.DefaultableBoolean.True
                .Header.Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
                mdiTKPC.uppdFindTKPC.ShowDialog()
            End With

        ElseIf whichFindPrintOption = 2 Then
            If mdiTKPC.findPeopleRowSelected = 0 Then
                MsgBox("데이터를 선택한 후 다시 시도하시기 바랍니다", Nothing, "주의")
                Exit Sub
            End If

            Dim row As UltraGridRow
            For Each row In mdiTKPC.ugFind.Rows
                If row.Cells("choose").Text = "True" Then
                    row.Selected = True
                    row.Appearance.BackColor = Color.White
                    mdiTKPC.ugFind.GetRowFromPrintRow(row)
                    row.Hidden = False
                Else
                    row.Selected = False
                    row.Hidden = True
                End If
            Next

            With mdiTKPC.ugpdFindTKPC
                .Header.TextCenter = txtTitle.Text
                .Header.Appearance.FontData.Underline = Infragistics.Win.DefaultableBoolean.True
                .Header.Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
                mdiTKPC.uppdFindTKPC.ShowDialog()
            End With

            For Each row In mdiTKPC.ugFind.Rows
                row.Hidden = False
                If row.Cells("choose").Text = "True" Then
                    row.Appearance.BackColor = Color.LightPink
                End If
            Next
        ElseIf whichFindPrintOption = 3 Then
            Dim row As UltraGridRow
            For Each row In mdiTKPC.ugFind.Rows
                If row.Cells("choose").Text = "False" Then
                    row.Selected = True
                    mdiTKPC.ugFind.GetRowFromPrintRow(row)
                    row.Hidden = False
                Else
                    row.Selected = False
                    row.Hidden = True
                End If
            Next

            With mdiTKPC.ugpdFindTKPC
                .Header.TextCenter = txtTitle.Text
                .Header.Appearance.FontData.Underline = Infragistics.Win.DefaultableBoolean.True
                .Header.Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
                mdiTKPC.uppdFindTKPC.ShowDialog()
            End With
            Me.Cursor = Cursors.Default

            For Each row In mdiTKPC.ugFind.Rows
                row.Hidden = False
            Next
        End If

        Me.Cursor = Cursors.Default
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub btCancel_Click(sender As Object, e As EventArgs) Handles btCancel.Click
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub txtTitle_KeyDown(sender As Object, e As KeyEventArgs) Handles txtTitle.KeyDown
        If e.KeyCode = Keys.Enter Then
            btOK_Click(sender, e)
        End If
    End Sub
End Class
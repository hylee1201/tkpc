Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports Infragistics.Win.UltraWinGrid

Public Class frmPerson
    Private findPeopleRowSelected As Integer = 0
    Private familyMode As Boolean = False
    Private countOfLoaded As Integer = 0

    Private Sub frmPerson_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        tscbDataOption.SelectedIndex = 1
    End Sub

    Private Sub fillDgFind()
        Dim personDAL As TKPC.DAL.PersonDAL = TKPC.DAL.PersonDAL.getInstance()
        Dim ds As DataSet = Nothing

        Try
            Me.Cursor = Cursors.WaitCursor

            'Dim personIds As String = ""
            'Dim cp As List(Of String) = readPersonIdFromDgChosen(0)
            'If cp.Count > 0 Then
            '    personIds = String.Join(",", cp)
            'End If

            ds = personDAL.getPeopleList(8, Nothing, Nothing, Nothing, tscbDataOption.SelectedIndex)
            m_Global.setNavigatorBinding(ds, 2, True, "등록교인")

        Catch ex As Exception
            Logger.LogError(ex.ToString)
        Finally
            personDAL = Nothing
            ds = Nothing
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    'Private Sub frmPerson_Activated(sender As Object, e As EventArgs) Handles Me.Activated
    '    Me.TopMost = True
    'End Sub

    Private Sub tsbtExit_Click(sender As Object, e As EventArgs) Handles tsbtExit.Click
        Me.Close()
        Me.Dispose()
    End Sub

    Private Function readPersonIdFromDgChosen() As List(Of String)
        readPersonIdFromDgChosen = New List(Of String)
        If ugResult.Rows.Count > 0 Then
            For i = 0 To ugResult.Rows.FilteredInRowCount - 1
                readPersonIdFromDgChosen.Add(CType(ugResult.Rows(i).Cells("spouses").Value, String))
            Next
        End If
    End Function

    Private Sub ugResult_KeyDown(sender As Object, e As KeyEventArgs) Handles ugResult.KeyDown
        If e.KeyCode = Keys.Down Then
            ugResult.ActiveRowScrollRegion.Scroll(Infragistics.Win.UltraWinGrid.RowScrollAction.LineDown)
        ElseIf e.KeyCode = Keys.Up Then
            ugResult.ActiveRowScrollRegion.Scroll(Infragistics.Win.UltraWinGrid.RowScrollAction.LineUp)
        ElseIf e.KeyCode = Keys.PageDown Then
            ugResult.ActiveRowScrollRegion.Scroll(Infragistics.Win.UltraWinGrid.RowScrollAction.PageDown)
        ElseIf e.KeyCode = Keys.PageUp Then
            ugResult.ActiveRowScrollRegion.Scroll(Infragistics.Win.UltraWinGrid.RowScrollAction.PageUp)
        End If
    End Sub

    Private Sub ugResult_InitializeLayout(sender As Object, e As InitializeLayoutEventArgs) Handles ugResult.InitializeLayout
        If familyMode = False Then
            Dim checkColumn As UltraGridColumn = e.Layout.Bands(0).Columns.Add("choose", "choose")
            checkColumn.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox
            checkColumn.DataType = Type.GetType("System.Boolean")
            checkColumn.DefaultCellValue = False
            checkColumn.CellActivation = Activation.AllowEdit
            checkColumn.Header.VisiblePosition = 0
            checkColumn.CellClickAction = CellClickAction.Edit
            ugResult.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True

            Dim btnColumn As UltraGridColumn = e.Layout.Bands(0).Columns.Add("picture", "picture")
            btnColumn.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button
            btnColumn.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always
            'btnColumn.CellButtonAppearance.Image = "C:\Users\hylee\Documents\TKPC\system\VBExpress\TKPCAdminReportMaker\TKPCAdminReportMaker\TKPCAdminReportMaker\Resources\view_picture.png"
            btnColumn.CellButtonAppearance.TextHAlign = Infragistics.Win.HAlign.Center
            btnColumn.Width = 50
            btnColumn.DefaultCellValue = False
            btnColumn.CellActivation = Activation.AllowEdit
            btnColumn.Header.VisiblePosition = 1
            btnColumn.CellClickAction = CellClickAction.Edit
        End If
    End Sub

    Private Sub tsbtCheckAll_Click(sender As Object, e As EventArgs) Handles tsbtCheckAll.Click
        With ugResult
            For i As Integer = 0 To .Rows.Count - 1
                .Rows(i).Cells("choose").Value = True
                .Rows(i).Appearance.BackColor = Color.LightPink
            Next
        End With
        findPeopleRowSelected = ugResult.Rows.Count
    End Sub

    Private Sub tsbtUncheckAll_Click(sender As Object, e As EventArgs) Handles tsbtUncheckAll.Click
        With ugResult
            For i As Integer = 0 To .Rows.Count - 1
                .Rows(i).Cells("choose").Value = False
                .Rows(i).Appearance.BackColor = Color.White
            Next
        End With
        findPeopleRowSelected = 0
    End Sub

    Private Sub ugResult_CellChange(sender As Object, e As CellEventArgs) Handles ugResult.CellChange
        If e.Cell.Text = "True" Then
            findPeopleRowSelected += 1
            e.Cell.Row.Appearance.BackColor = Color.LightPink
            e.Cell.Row.Selected = True
        Else
            findPeopleRowSelected -= 1
            e.Cell.Row.Appearance.BackColor = Color.White
            e.Cell.Row.Selected = False
        End If
    End Sub

    Private Sub 모든데이터ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 모든데이터ToolStripMenuItem.Click
        Me.Cursor = Cursors.WaitCursor
        Try
            Dim theFile As String = ROOT_FOLDER + "\" + m_Constant.TEMP_FILE 'Application.StartupPath + "\\" + BOOK_LIST_EXCEL
            Me.ugexcel.Export(Me.ugResult, theFile)
            Diagnostics.Process.Start(theFile)

        Catch ex As Exception
            MessageBox.Show(ex.Message + " If the file called tkpc_member_list.xls at " + ROOT_FOLDER + " is open, please close it and try again.")
            Logger.LogError(ex.ToString)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub 선택한데이터만ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 선택한데이터만ToolStripMenuItem.Click
        If findPeopleRowSelected = 0 Then
            MsgBox("데이터를 선택한 후 다시 시도하시기 바랍니다")
            Exit Sub
        End If

        Dim row As UltraGridRow
        For Each row In ugResult.Rows
            If row.Cells("choose").Text = "True" Then
                row.Selected = True
                ugResult.GetRowFromPrintRow(row)
                row.Hidden = False
            Else
                row.Selected = False
                row.Hidden = True
            End If
        Next

        Me.Cursor = Cursors.WaitCursor
        Try
            Dim theFile As String = ROOT_FOLDER + "\" + m_Constant.TEMP_FILE 'Application.StartupPath + "\\" + BOOK_LIST_EXCEL
            Me.ugexcel.Export(Me.ugResult, theFile)
            Diagnostics.Process.Start(theFile)

        Catch ex As Exception
            MessageBox.Show(ex.Message + " If the file called tkpc_member_list.xls at " + ROOT_FOLDER + " is open, please close it and try again.")
            Logger.LogError(ex.ToString)
        Finally
            Me.Cursor = Cursors.Default
        End Try

        For Each row In ugResult.Rows
            row.Hidden = False
        Next

    End Sub

    Private Sub 선택하지않은데이터만ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 선택하지않은데이터만ToolStripMenuItem.Click
        Dim row As UltraGridRow
        For Each row In ugResult.Rows
            If row.Cells("choose").Text = "False" Then
                row.Selected = True
                ugResult.GetRowFromPrintRow(row)
                row.Hidden = False
            Else
                row.Selected = False
                row.Hidden = True
            End If
        Next

        Me.Cursor = Cursors.WaitCursor
        Try
            Dim theFile As String = ROOT_FOLDER + "\" + m_Constant.TEMP_FILE 'Application.StartupPath + "\\" + BOOK_LIST_EXCEL
            Me.ugexcel.Export(Me.ugResult, theFile)
            Diagnostics.Process.Start(theFile)

        Catch ex As Exception
            MessageBox.Show(ex.Message + " If the file called tkpc_member_list.xls at " + ROOT_FOLDER + " is open, please close it and try again.")
            Logger.LogError(ex.ToString)
        Finally
            Me.Cursor = Cursors.Default
        End Try

        For Each row In ugResult.Rows
            row.Hidden = False
        Next
    End Sub

    Private Sub 입교로ToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Me.TopMost = False
        frmConversion2.lblOptionTitle.Text = "입교한 교인으로 처리합니다."
        frmConversion2.ShowDialog()
    End Sub

    Private Sub 세례로ToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Me.TopMost = False
        frmConversion2.lblOptionTitle.Text = "세례교인으로 처리합니다."
        frmConversion2.ShowDialog()
    End Sub

    Private Sub 유아세례로ToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Me.TopMost = False
        frmConversion2.lblOptionTitle.Text = "유아세례교인으로 처리합니다."
        frmConversion2.ShowDialog()
    End Sub

    Private Sub 직분변경ToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Me.TopMost = False
        frmConversion.lblOptionTitle.Text = "직분변경을 합니다."
        frmConversion.ShowDialog()
    End Sub

    Private Sub writeSurveyForm()
        m_Global.writeSurveyForm(1)
    End Sub

    Private Sub ugResult_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles ugResult.InitializeRow
        If familyMode = True Then
            If e.Row.Cells("lvl").Text = m_Constant.FAMILY_LEVEL_1 Then
                e.Row.CellAppearance.BackColor = Color.LightGreen
            ElseIf e.Row.Cells("lvl").Text = m_Constant.FAMILY_LEVEL_0 Then
                e.Row.CellAppearance.BackColor = Color.LightSalmon
            Else
                e.Row.CellAppearance.BackColor = Color.White
            End If
        Else
            Dim fileName As String = e.Row.Cells("photo_file_name").Text
            If String.IsNullOrEmpty(fileName) = False Then
                If System.IO.File.Exists(m_Constant.PERSON_IMAGE_FOLDER + fileName) Then
                    e.Row.Cells("picture").Value = "view"
                    e.Row.Cells("picture").ButtonAppearance.BackColor = Color.LightGreen
                Else
                    e.Row.Cells("picture").Value = "download"
                End If
            Else
                e.Row.Cells("picture").Value = "upload"
            End If
        End If
    End Sub

    Private Sub tscbDataOption_SelectedIndexChanged(sender As Object, e As EventArgs) Handles tscbDataOption.SelectedIndexChanged
        familyMode = False
        fillDgFind()
        countOfLoaded = 1
    End Sub

    Private Sub 교적카드로ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 교적카드로ToolStripMenuItem.Click

    End Sub

    Private Sub 설문지양식으로ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 설문지양식으로ToolStripMenuItem.Click
        If familyMode = False Then
            m_Global.getFamilyList(1, readPersonIdFromDgChosen())
        End If
        writeSurveyForm()
    End Sub

    Private Sub 교인주소록으로ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 교인주소록으로ToolStripMenuItem.Click

    End Sub

    Private Sub 개인모드로ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 개인모드로ToolStripMenuItem.Click
        If familyMode = True Then
            familyMode = False
            fillDgFind()
            countOfLoaded = 1
        Else
            MsgBox("이미 개인모드로 변경되었습니다.")
        End If
    End Sub

    Private Sub 가족모드로ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 가족모드로ToolStripMenuItem.Click
        If familyMode = False Then
            familyMode = True
            m_Global.getFamilyList(1, readPersonIdFromDgChosen())
        Else
            MsgBox("이미 가족모드로 변경되었습니다.")
        End If
    End Sub

    Private Sub ugResult_ClickCellButton(sender As Object, e As CellEventArgs) Handles ugResult.ClickCellButton
        With frmDetail
            .personId = e.Cell.Row.Cells("person_id").Text

            If e.Cell.Row.Cells("picture").Text = "view" Then
                .photo_file_name = m_Constant.PERSON_IMAGE_FOLDER + e.Cell.Row.Cells("photo_file_name").Text
            Else
                .photo_file_name = m_Constant.PERSON_IMAGE_FOLDER + "no_image.jpg"
            End If
            .ShowDialog()
        End With
    End Sub
End Class
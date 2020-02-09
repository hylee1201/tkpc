Imports Infragistics.Win.UltraWinGrid
Imports System.Collections
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Text

Public Class frmVisitInfoOption
    Public cp As List(Of String) = Nothing
    Public visitOptionPersonTable As Hashtable = New Hashtable
    Dim personDAL As TKPC.DAL.PersonDAL = TKPC.DAL.PersonDAL.getInstance()

    Private Sub frmVisitInfoOption_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim personIds As String = ""
        If cp.Count > 0 Then
            personIds = String.Join(",", cp)
            mdiTKPC.prevFoundPersonList = cp
            Dim ds As DataSet = Nothing
            Dim personDAL As TKPC.DAL.PersonDAL = TKPC.DAL.PersonDAL.getInstance()
            ds = personDAL.getPeopleList(8, personIds, "", "", 0)
            m_Global.setNavigatorBinding(ds, 3, True, "")
        End If
    End Sub

    Private Sub ugResult_InitializeLayout(sender As Object, e As InitializeLayoutEventArgs) Handles ugResult.InitializeLayout
        Dim checkColumn As UltraGridColumn = e.Layout.Bands(0).Columns.Add("choose", "choose")
        checkColumn.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox
        checkColumn.DataType = Type.GetType("System.Boolean")
        checkColumn.DefaultCellValue = False
        checkColumn.CellActivation = Activation.AllowEdit
        checkColumn.Header.VisiblePosition = 0
        checkColumn.CellClickAction = CellClickAction.Edit
        checkColumn.Width = 50

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

        With ugResult
            .DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True
            .DisplayLayout.Bands(0).Columns(0).Width = 60 'person_id
            .DisplayLayout.Bands(0).Columns(0).Hidden = True 'person_id
            .DisplayLayout.Bands(0).Columns(1).Width = 70 'Korean Name
            .DisplayLayout.Bands(0).Columns(2).Width = 50 'title
            .DisplayLayout.Bands(0).Columns(3).Width = 70 'Last Name
            .DisplayLayout.Bands(0).Columns(4).Width = 100 'First Name
            .DisplayLayout.Bands(0).Columns(5).Width = 30 'Gender
            .DisplayLayout.Bands(0).Columns(6).Width = 70 'DOB
            .DisplayLayout.Bands(0).Columns(7).Width = 35 'Age
            .DisplayLayout.Bands(0).Columns(8).Width = 80 'Cell Phone
            .DisplayLayout.Bands(0).Columns(9).Width = 80 'home Phone
            .DisplayLayout.Bands(0).Columns(10).Width = 80 'work Phone
            .DisplayLayout.Bands(0).Columns(11).Width = 295 'Home address
            .DisplayLayout.Bands(0).Columns(12).Width = 50 'status
            .DisplayLayout.Bands(0).Columns(13).Width = 90 'status_date
            .DisplayLayout.Bands(0).Columns(14).Width = 90 'deceased_on
            .DisplayLayout.Bands(0).Columns(15).Width = 80 'spouses
            .DisplayLayout.Bands(0).Columns(15).Hidden = True
            .DisplayLayout.Bands(0).Columns(16).Hidden = True 'photo_file_name
            .DisplayLayout.Override.RowAppearance.TextVAlign = Infragistics.Win.VAlign.Middle
        End With
    End Sub

    Private Sub btOK_Click(sender As Object, e As EventArgs) Handles btOK.Click
        writeMembershipCard()
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub writeMembershipCard()
        With ugResult
            For i As Integer = 0 To .Rows.Count - 1
                If .Rows(i).Cells("choose").Text = "True" Then
                    Dim personId As String = .Rows(i).Cells("person_id").Text
                    If visitOptionPersonTable.ContainsKey(personId) = False Then
                        visitOptionPersonTable.Add(personId, personId)
                    End If
                End If
            Next
        End With

        If ugResult.Rows.Count > 0 Then
            m_Global.getFamilyList(0, cp)
        End If
        m_Global.writeMembershipCard(0, visitOptionPersonTable, False)
        mdiTKPC.prevFoundPersonList = cp
    End Sub

    Private Sub tsbtFindExit_Click(sender As Object, e As EventArgs) Handles tsbtFindExit.Click
        Me.Dispose()
        Me.Close()
    End Sub

    Private Sub btCancel_Click(sender As Object, e As EventArgs) Handles btCancel.Click
        Me.Dispose()
        Me.Close()
    End Sub

    Private Sub tsbtCheckAll_Click(sender As Object, e As EventArgs) Handles tsbtCheckAll.Click
        With ugResult
            For i As Integer = 0 To .Rows.Count - 1
                .Rows(i).Cells("choose").Value = True
                .Rows(i).Appearance.BackColor = Color.LightPink
            Next
        End With
    End Sub

    Private Sub tsbtUncheckAll_Click(sender As Object, e As EventArgs) Handles tsbtUncheckAll.Click
        With ugResult
            For i As Integer = 0 To .Rows.Count - 1
                .Rows(i).Cells("choose").Value = False
                .Rows(i).Appearance.BackColor = Color.White
            Next
        End With
    End Sub

    Private Sub ugResult_CellChange(sender As Object, e As CellEventArgs) Handles ugResult.CellChange
        If e.Cell.Text = "True" Then
            e.Cell.Row.Appearance.BackColor = Color.LightPink
        Else
            e.Cell.Row.Appearance.BackColor = Color.White
        End If
    End Sub

    Private Sub ugResult_AfterRowActivate(sender As Object, e As EventArgs) Handles ugResult.AfterRowActivate
        Dim activeRow As UltraGridRow = ugResult.ActiveRow

        If activeRow.Band Is ugResult.DisplayLayout.Bands(0) Then
            Dim visitList As List(Of TKPC.Entity.VisitEnt) = personDAL.getVisitedFamily(activeRow.Cells("spouses").Text)
            If visitList IsNot Nothing Then
                With dgVisitInfo
                    .Rows.Clear()
                    For i As Integer = 0 To visitList.Count - 1
                        Dim visitEnt As TKPC.Entity.VisitEnt = visitList.Item(i)
                        If visitEnt IsNot Nothing Then
                            .Rows.Add(m_Global.formatDate(0, visitEnt.visitDate, ""), visitEnt.visitedKoreanName, visitEnt.note)
                        End If
                    Next
                End With
            End If
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

    Private Sub ugResult_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles ugResult.InitializeRow
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

        'Dim deceased As String = e.Row.Cells("deceased_on").Text
        'If String.IsNullOrEmpty(deceased) = False Then
        '    e.Row.Cells("status").Value = "고인"
        'End If
    End Sub
End Class
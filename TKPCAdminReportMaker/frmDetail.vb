Imports System.Collections
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Text
Imports Infragistics.Win.UltraWinGrid

Public Class frmDetail
    Public personId As String = ""
    Public photo_file_name As String = ""
    Dim visitList As List(Of TKPC.Entity.VisitEnt) = Nothing
    Dim personIds As String = ""
    Dim familyLevelTable As Hashtable = New Hashtable
    Dim commentList As List(Of TKPC.Entity.CommentEnt) = Nothing
    Dim isFirst As Boolean = False

    Private Sub frmDetail_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim personList As List(Of String) = New List(Of String)
        personList.Add(personId)

        m_Global.getFamilyList(1, personList)
        pbDetail.ImageLocation = photo_file_name

        If ugResult.Rows.Count > 0 Then
            tsbtExcel.Visible = True
            tsbtSurvey.Visible = True
        End If

        For i = 0 To ugResult.Rows.Count - 1
            personIds += ugResult.Rows(i).Cells("id").Text
            If (i < ugResult.Rows.Count - 1) Then
                personIds += ","
            End If
        Next

        Dim personDAL As TKPC.DAL.PersonDAL = TKPC.DAL.PersonDAL.getInstance()
        visitList = personDAL.getVisitedFamily(personIds)
        If visitList IsNot Nothing And visitList.Count > 0 Then
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

        If ugResult.Rows.Count > 0 Then
            commentList = personDAL.getComments(personIds)
            If commentList IsNot Nothing Then
                With dgComment
                    .Rows.Clear()
                    For i As Integer = 0 To commentList.Count - 1
                        Dim commentEnt As TKPC.Entity.CommentEnt = commentList.Item(i)
                        If commentEnt IsNot Nothing Then
                            Dim lvl As String = familyLevelTable.Item(commentEnt.personId).ToString
                            .Rows.Add(commentEnt.personId, lvl, commentEnt.personName, commentEnt.authorName, commentEnt.body)
                        End If
                    Next
                    .Sort(dgComment.Columns("가족구분3"), System.ComponentModel.ListSortDirection.Ascending)
                End With
            End If
        End If

        If m_Global.accessLevel = m_Constant.ACCESS_LEVEL_OFFICE Then
            tcDetails.TabPages.Remove(TabPage2)
            tsbtCard.Visible = False
            tsbtCard2.Visible = True
        Else
            tcDetails.TabPages.Remove(TabPage4)
            If visitList.Count > 0 Then
                tsbtCard.Visible = True
                tsbtCard2.Visible = False
                심방기록과함께ToolStripMenuItem.Visible = True
                심방기록없이출력ToolStripMenuItem.Visible = True
            Else
                tsbtCard.Visible = False
                tsbtCard2.Visible = True
            End If
        End If
    End Sub

    Private Sub tsbtExit_Click(sender As Object, e As EventArgs) Handles tsbtExit.Click
        Me.Dispose()
        Me.Close()
    End Sub

    Private Sub tsbtExcel_Click(sender As Object, e As EventArgs) Handles tsbtExcel.Click
        Me.Cursor = Cursors.WaitCursor
        Try
            Dim theFile As String = FILE_FOLDER + "\" + m_Constant.TEMP_FILE 'Application.StartupPath + "\\" + BOOK_LIST_EXCEL
            Dim fileHelper As New FileHelper

            Me.ugexcel.Export(Me.ugResult, theFile)
            Diagnostics.Process.Start(theFile)


        Catch ex As Exception
            MessageBox.Show(ex.Message + " If the file called temp.xls at " + FILE_FOLDER + " is open, please close it and try again.")
            Logger.LogError(ex.ToString)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub tsbtSurvey_Click(sender As Object, e As EventArgs) Handles tsbtSurvey.Click
        m_Global.writeSurveyForm(2, 1, 0)
    End Sub

    Private Sub tcDetails_SelectedIndexChanged(sender As Object, e As EventArgs) Handles tcDetails.SelectedIndexChanged
        Dim personDAL As TKPC.DAL.PersonDAL = TKPC.DAL.PersonDAL.getInstance()
        Dim offeringDAL As TKPC.DAL.OfferingDAL = TKPC.DAL.OfferingDAL.getInstance()

        Try
            If tcDetails.SelectedTab.Text = "심방기록" Then
                If ugResult.Rows.Count > 0 Then
                    tsbtSurvey.Visible = True
                    tsbtExcel.Visible = True
                End If
            ElseIf tcDetails.SelectedTab.Text = "마일스톤" Then
                If ugResult.Rows.Count > 0 Then
                    tsbtSurvey.Visible = True
                    tsbtExcel.Visible = True

                    Dim stoneList As List(Of TKPC.Entity.MilestoneEnt) = personDAL.getMilestones(personIds)
                    If stoneList IsNot Nothing Then
                        With dgMilestone
                            .Rows.Clear()
                            For i As Integer = 0 To stoneList.Count - 1
                                Dim stoneEnt As TKPC.Entity.MilestoneEnt = stoneList.Item(i)
                                If stoneEnt IsNot Nothing Then
                                    Dim lvl As String = familyLevelTable.Item(stoneEnt.personId).ToString
                                    .Rows.Add(stoneEnt.personId, lvl, stoneEnt.koreanName, m_Global.formatDate(0, stoneEnt.effectiveDate, ""), stoneEnt.type, stoneEnt.name)
                                End If
                            Next
                            .Sort(dgMilestone.Columns("가족구분"), System.ComponentModel.ListSortDirection.Ascending)
                        End With
                    End If
                End If
            ElseIf tcDetails.SelectedTab.Text = "헌금봉투와 헌금영수증" Then
                If ugResult.Rows.Count > 0 Then
                    Dim ds As DataSet = offeringDAL.getOfferingForDetails(personIds)
                    With ugOffering
                        .DataSource = ds
                    End With
                End If
            ElseIf tcDetails.SelectedTab.Text = "주소" Then
                If ugResult.Rows.Count > 0 Then
                    Dim addressList As List(Of TKPC.Entity.PersonEnt) = personDAL.getAddresses(personIds)
                    If addressList IsNot Nothing Then
                        With dgAddress
                            .Rows.Clear()
                            For i As Integer = 0 To addressList.Count - 1
                                Dim personEnt As TKPC.Entity.PersonEnt = addressList.Item(i)
                                If personEnt IsNot Nothing Then
                                    Dim lvl As String = familyLevelTable.Item(personEnt.personId).ToString
                                    .Rows.Add(personEnt.personId, lvl, personEnt.koreanName, personEnt.addressType, personEnt.address)
                                End If
                            Next
                            .Sort(dgAddress.Columns("가족구분2"), System.ComponentModel.ListSortDirection.Ascending)
                        End With

                        If isFirst = False Then
                            Dim btn As New DataGridViewButtonColumn()
                            dgAddress.Columns.Add(btn)
                            btn.HeaderText = "Google Map"
                            btn.Text = "View"
                            btn.Name = "Map"
                            btn.UseColumnTextForButtonValue = True

                            Dim btn2 As New DataGridViewButtonColumn()
                            dgAddress.Columns.Add(btn2)
                            btn2.HeaderText = "From TKPC"
                            btn2.Text = "View"
                            btn2.Name = "Dir"
                            btn2.UseColumnTextForButtonValue = True
                            isFirst = True
                        End If
                    End If
                End If
                'Else
                '    If ugResult.Rows.Count > 0 Then
                '        Dim commentList As List(Of TKPC.Entity.CommentEnt) = personDAL.getComments(personIds)
                '        If commentList IsNot Nothing Then
                '            With dgComment
                '                .Rows.Clear()
                '                For i As Integer = 0 To commentList.Count - 1
                '                    Dim commentEnt As TKPC.Entity.CommentEnt = commentList.Item(i)
                '                    If commentEnt IsNot Nothing Then
                '                        Dim lvl As String = familyLevelTable.Item(commentEnt.personId).ToString
                '                        .Rows.Add(commentEnt.personId, lvl, commentEnt.personName, commentEnt.authorName, commentEnt.body)
                '                    End If
                '                Next
                '                .Sort(dgComment.Columns("가족구분3"), System.ComponentModel.ListSortDirection.Ascending)
                '            End With
                '        End If
                '    End If
            End If

        Catch ex As Exception
            Logger.LogError(ex.ToString)
        Finally
            personDAL = Nothing
            offeringDAL = Nothing
        End Try
    End Sub

    Private Sub 심방기록과함께ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 심방기록과함께ToolStripMenuItem.Click
        Dim visitOptionPersonTable As Hashtable = New Hashtable
        With ugResult
            For i As Integer = 0 To .Rows.Count - 1
                Dim personId As String = .Rows(i).Cells("id").Text
                If visitOptionPersonTable.ContainsKey(personId) = False Then
                    visitOptionPersonTable.Add(personId, personId)
                End If
            Next
        End With

        If commentList IsNot Nothing And commentList.Count > 0 Then
            Dim result As Integer = MessageBox.Show("비고내용이 있습니다. 같이 출력하시겠습니까?", "선택", MessageBoxButtons.YesNo)
            If result = DialogResult.Yes Then
                m_Global.writeMembershipCard(1, visitOptionPersonTable, True)
            Else
                m_Global.writeMembershipCard(1, visitOptionPersonTable, False)
            End If
        Else
            m_Global.writeMembershipCard(1, visitOptionPersonTable, False)
        End If
    End Sub

    Private Sub 심방기록없이출력ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 심방기록없이출력ToolStripMenuItem.Click
        If commentList IsNot Nothing And commentList.Count > 0 Then
            Dim result As Integer = MessageBox.Show("비고내용이 있습니다. 같이 출력하시겠습니까?", "선택", MessageBoxButtons.YesNo)
            If result = DialogResult.Yes Then
                m_Global.writeMembershipCard(1, Nothing, True)
            Else
                m_Global.writeMembershipCard(1, Nothing, False)
            End If
        Else
            m_Global.writeMembershipCard(1, Nothing, False)
        End If
    End Sub

    Private Sub ugResult_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles ugResult.InitializeRow
        Dim type As String = e.Row.Cells("type").Text
        Dim id As String = e.Row.Cells("id").Text

        If type = "세대주" Then
            type = "1-" + type
        ElseIf type = "배우자" Then
            type = "2-" + type
        Else
            type = "3-" + type
        End If

        If familyLevelTable.ContainsKey(id) = False Then
            familyLevelTable.Add(id, type)
        End If

    End Sub

    Private Sub ugResult_InitializeLayout(sender As Object, e As InitializeLayoutEventArgs) Handles ugResult.InitializeLayout
        Dim checkColumn As UltraGridColumn = e.Layout.Bands(0).Columns.Add("제외", "제외")
        checkColumn.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox
        checkColumn.DataType = Type.GetType("System.Boolean")
        checkColumn.DefaultCellValue = False
        checkColumn.CellActivation = Activation.AllowEdit
        checkColumn.Header.VisiblePosition = 0
        checkColumn.CellClickAction = CellClickAction.Edit
        checkColumn.Width = 50

        e.Layout.Override.HeaderClickAction = HeaderClickAction.Select

        e.Layout.Bands(0).Columns("PhotoName").Hidden = True
        e.Layout.Bands(0).Columns("id").Hidden = True
    End Sub

    Private Sub tsbtCard2_Click(sender As Object, e As EventArgs) Handles tsbtCard2.Click
        If commentList IsNot Nothing And commentList.Count > 0 Then
            Dim result As Integer = MessageBox.Show("비고내용이 있습니다. 같이 출력하시겠습니까?", "선택", MessageBoxButtons.YesNo)
            If result = DialogResult.Yes Then
                m_Global.writeMembershipCard(1, Nothing, True)
            Else
                m_Global.writeMembershipCard(1, Nothing, False)
            End If
        Else
            m_Global.writeMembershipCard(1, Nothing, False)
        End If
    End Sub

    Private Sub ugOffering_InitializeLayout(sender As Object, e As InitializeLayoutEventArgs) Handles ugOffering.InitializeLayout
        Dim btnColumn As UltraGridColumn = e.Layout.Bands(0).Columns.Add("receipt", "receipt")
        btnColumn.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button
        btnColumn.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always
        'btnColumn.CellButtonAppearance.Image = "C:\Users\hylee\Documents\TKPC\system\VBExpress\TKPCAdminReportMaker\TKPCAdminReportMaker\TKPCAdminReportMaker\Resources\view_picture.png"
        btnColumn.CellButtonAppearance.TextHAlign = Infragistics.Win.HAlign.Center
        btnColumn.Width = 60
        btnColumn.DefaultCellValue = False
        btnColumn.CellActivation = Activation.AllowEdit
        btnColumn.Header.VisiblePosition = 0
        btnColumn.CellClickAction = CellClickAction.Edit

        'id, e.year, e.offering_number, p.korean_name, UPPER(p.last_name) last_name, UPPER(p.first_name) first_name, " + vbCrLf)
        'INITCAP(a.street) street, INITCAP(a.city) city, INITCAP(a.province) province, UPPER(a.postal_code) postal_code " + vbCrLf)

        With ugOffering
            .DisplayLayout.Bands(0).Columns(0).Width = 60 'person_id
            .DisplayLayout.Bands(0).Columns(0).Hidden = True 'person_id
            .DisplayLayout.Bands(0).Columns(1).Width = 50 'year
            .DisplayLayout.Bands(0).Columns(2).Width = 60 'number
            .DisplayLayout.Bands(0).Columns(3).Width = 90 'korean_name
            .DisplayLayout.Bands(0).Columns(4).Width = 80 'last_name
            .DisplayLayout.Bands(0).Columns(5).Width = 80 'first_name
            .DisplayLayout.Bands(0).Columns(6).Width = 200 'street
            .DisplayLayout.Bands(0).Columns(7).Width = 70 'city
            .DisplayLayout.Bands(0).Columns(8).Width = 70 'province
            .DisplayLayout.Bands(0).Columns(9).Width = 60 'postal code
        End With
    End Sub

    Private Sub ugOffering_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles ugOffering.InitializeRow
        If CType(e.Row.Cells("year").Value, Integer) < 2019 Then
            e.Row.Cells("receipt").Hidden = CBool(Infragistics.Win.DefaultableBoolean.True)
        Else
            e.Row.Cells("receipt").Value = "발행"
        End If
    End Sub

    Private Sub ugOffering_ClickCellButton(sender As Object, e As CellEventArgs) Handles ugOffering.ClickCellButton
        With frmReceiptInfo
            .flag = 1
            .offeringYear = e.Cell.Row.Cells("year").Text
            .personId = e.Cell.Row.Cells("id").Text
            .lastName = e.Cell.Row.Cells("last_name").Text
            .firstName = e.Cell.Row.Cells("first_name").Text
            .street = e.Cell.Row.Cells("street").Text
            .city = e.Cell.Row.Cells("city").Text
            .postalcode = e.Cell.Row.Cells("postal_code").Text
            .offeringNumber = e.Cell.Row.Cells("number").Text
            .ShowDialog()
        End With
    End Sub

    Private Sub dgAddress_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgAddress.CellClick
        If e.ColumnIndex = 5 Then
            Dim address As String = dgAddress.Rows(e.RowIndex).Cells(e.ColumnIndex - 1).Value.ToString
            If String.IsNullOrEmpty(address) = False Then
                Dim url As String = "https://www.google.ca/maps/place/" + address.Replace(" ", "+")
                'url = "https://www.google.ca/maps/place/79+Wellington+St+W,+Toronto,+ON+M5K+1J5"
                System.Diagnostics.Process.Start(url)
            Else
                MsgBox("주소를 업데이트하시기 바랍니다.", Nothing, "요망")
            End If
        ElseIf e.ColumnIndex = 6 Then '교회에서 가는 길
            Dim address As String = dgAddress.Rows(e.RowIndex).Cells(e.ColumnIndex - 2).Value.ToString
            If String.IsNullOrEmpty(address) = False Then
                Dim url As String = "https://www.google.ca/maps/dir/" + m_Constant.ADDRESS_TKPC.Replace(" ", "+") + "/" + address.Replace(" ", "+")
                'https://www.google.ca/maps/dir/62+Ian+Baron+Ave,+Unionville,+ON+L3R+4R2/79+Wellington+St+W,+Toronto,+ON+M5K+1J5/
                System.Diagnostics.Process.Start(url)
            Else
                MsgBox("주소를 업데이트하시기 바랍니다.", Nothing, "요망")
            End If
        End If
    End Sub
End Class
Imports System.Collections

Public Class frmUpdateAllOption
    Public dataHT As Hashtable = New Hashtable
    Public milestoneType As String = ""

    Private Sub btOK_Click(sender As Object, e As EventArgs) Handles btOK.Click
        Dim personDAL As TKPC.DAL.PersonDAL = TKPC.DAL.PersonDAL.getInstance()
        Dim pickedDate As DateTime = dtpOption.Value

        Try
            If milestoneType = m_Constant.MILESTONE_TITLE And cbOption.Text = m_Constant.MILESTONE_NAME_SEORI Then
                If pickedDate.DayOfYear <> 1 Then
                    MessageBox.Show("서리집사의 일괄처리 날짜를 1월1일로 맞추어주시기 바랍니다.")
                    Exit Sub
                End If
            End If

            Me.Cursor = Cursors.WaitCursor
            If dgChosen.RowCount > 0 Then
                Dim cntOfUpdate As Integer = 0
                Dim cntOfInsert As Integer = 0
                Dim success As Integer = 0

                pbAll.Value = 0
                pbAll.Visible = True
                pbAll.Maximum = dgChosen.RowCount

                For i = 0 To dgChosen.RowCount - 1
                    pbAll.Value += 1
                    Dim personId As String = CType(dgChosen.Rows(i).Cells("person_id").Value, String)
                    If personDAL.isMileStoneThere(personId, milestoneType, cbOption.Text, dtpOption.Text) = True Then
                        success = personDAL.updateMilestone(personId, milestoneType, cbOption.Text, dtpOption.Text)
                        If success = 1 Then
                            cntOfUpdate += 1
                        End If
                    Else
                        If String.IsNullOrEmpty(personId) = False Then
                            success = personDAL.insertMilestone(personId, milestoneType, cbOption.Text, dtpOption.Text)
                            If success = 1 Then
                                cntOfInsert += 1
                            End If
                        End If
                    End If
                Next

                pbAll.Visible = False
                MessageBox.Show(CType(cntOfUpdate, String) + "개의 레코드가 업데이트되었고 " + CType(cntOfInsert, String) + "개의 레코드가 추가되었습니다.")

            End If

        Catch ex As Exception
            Logger.LogError(ex.ToString)
        Finally
            Me.Cursor = Cursors.Default
            personDAL = Nothing
        End Try
    End Sub

    Private Sub btCancel_Click(sender As Object, e As EventArgs) Handles btCancel.Click
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub frmUpdateAllOption_Load(sender As Object, e As EventArgs) Handles Me.Load
        If dataHT IsNot Nothing And dataHT.Count > 0 Then

            If milestoneType = m_Constant.MILESTONE_TITLE Then
                cbOption.Visible = True
                cbOption.SelectedIndex = 3
                cbOption2.Visible = False
            Else
                cbOption2.Visible = True
                cbOption2.SelectedIndex = 0
                cbOption.Visible = False
            End If

            For Each personObj As DictionaryEntry In dataHT
                Dim personEnt As TKPC.Entity.PersonEnt = CType(personObj.Value, TKPC.Entity.PersonEnt)
                With dgChosen
                    If personEnt IsNot Nothing Then
                        .Rows.Add(personEnt.koreanName, personEnt.personId, personEnt.status, personEnt.age, personEnt.gender)
                    End If
                End With
                personEnt = Nothing
            Next
        End If
    End Sub
End Class
Imports System.Collections
Imports System.Collections.Generic

Public Class frmUpdateAllOption2
    Public dataHT As Hashtable = New Hashtable

    Private Sub dgChosen_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles dgChosen.CellEnter
        Dim personId As String = dgChosen.Rows(e.RowIndex).Cells(1).Value.ToString
        Dim personList As List(Of String) = New List(Of String)
        personList.Add(personId)

        m_Global.getFamilyList(2, personList)
    End Sub

    Private Sub frmUpdateAllOption2_Load(sender As Object, e As EventArgs) Handles Me.Load
        For Each personObj As DictionaryEntry In dataHT
            Dim personEnt As TKPC.Entity.PersonEnt = CType(personObj.Value, TKPC.Entity.PersonEnt)
            With dgChosen
                If personEnt IsNot Nothing Then
                    .Rows.Add(personEnt.koreanName, personEnt.personId, personEnt.status, personEnt.age, personEnt.gender)
                End If
            End With
            personEnt = Nothing
        Next
    End Sub

    Private Sub btCancel_Click(sender As Object, e As EventArgs) Handles btCancel.Click
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub tsbtExit_Click(sender As Object, e As EventArgs) Handles tsbtExit.Click
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub btOK_Click(sender As Object, e As EventArgs) Handles btOK.Click
        If MsgBox("선택한 사람들을 각 가족에게서 독립시킵니다. 독립후에는 가족리스트에서 볼 수 없습니다." + vbCrLf + "진행하시겠습니까? ", MsgBoxStyle.YesNo, "선택") = vbYes Then
            Dim personDAL As TKPC.DAL.PersonDAL = TKPC.DAL.PersonDAL.getInstance()

        End If
    End Sub
End Class
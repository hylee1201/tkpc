Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.Drawing
Imports System.IO
Imports System.Text
Imports System.Windows.Forms
Imports Infragistics.Win.UltraWinGrid

Public Class mdiTKPC
    Private rowIndex As Integer = 0
    Public findPeopleRowSelected As Integer = 0
    Private findPeopleHashTable As Hashtable = New Hashtable()
    Private familyMode As Boolean = False
    'frmVisitInfoOption uses this variable too
    Public prevFoundPersonList As List(Of String) = New List(Of String)

    Private Sub ShowNewForm(ByVal sender As Object, ByVal e As EventArgs)
        ' Create a new instance of the child form.
        Dim ChildForm As New System.Windows.Forms.Form
        ' Make it a child of this MDI form before showing it.
        ChildForm.MdiParent = Me

        m_ChildFormNumber += 1
        ChildForm.Text = "Window " & m_ChildFormNumber

        ChildForm.Show()
    End Sub

    Private Sub OpenFile(ByVal sender As Object, ByVal e As EventArgs)
        Dim OpenFileDialog As New OpenFileDialog
        OpenFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        OpenFileDialog.Filter = "Text Files (*.xls)|*.xls|All Files (*.*)|*.*"
        If (OpenFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK) Then
            Dim FileName As String = OpenFileDialog.FileName
            ' TODO: Add code here to open the file.
        End If
    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim SaveFileDialog As New SaveFileDialog
        SaveFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        SaveFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"

        If (SaveFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK) Then
            Dim FileName As String = SaveFileDialog.FileName
            ' TODO: Add code here to save the current contents of the form to a file.
        End If
    End Sub


    Private Sub ExitToolsStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub CutToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        ' Use My.Computer.Clipboard to insert the selected text or images into the clipboard
    End Sub

    Private Sub CopyToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        ' Use My.Computer.Clipboard to insert the selected text or images into the clipboard
    End Sub

    Private Sub PasteToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        'Use My.Computer.Clipboard.GetText() or My.Computer.Clipboard.GetData to retrieve information from the clipboard.
    End Sub

    Private Sub CascadeToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.Cascade)
    End Sub

    Private Sub TileVerticalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.TileVertical)
    End Sub

    Private Sub TileHorizontalToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.TileHorizontal)
    End Sub

    Private Sub ArrangeIconsToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        Me.LayoutMdi(MdiLayout.ArrangeIcons)
    End Sub

    Private Sub CloseAllToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs)
        ' Close all child forms of the parent.
        For Each ChildForm As Form In Me.MdiChildren
            ChildForm.Close()
        Next
    End Sub

    Private m_ChildFormNumber As Integer

    Private Sub mdiTKPC_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'cbCurrentEnableWhich(False, False)
        cbType.SelectedIndex = 0


        initDgChosen()
        Dim dLastSunday As Date = Now.AddDays(-(Now.DayOfWeek))
        dtpCurrentStartDate.Value = dLastSunday
        dtpCurrentEndDate.Value = dLastSunday
        cbFindWhat.SelectedIndex = 0
        chkType.Checked = True
        chkType.Enabled = True

        If whichIME = 0 Then
            whichIME = SwitchIME(hKrLayoutId, KbdKr)
        End If

        If m_Global.accessLevel = m_Constant.ACCESS_LEVEL_PASTOR Then
            cbPastor.SelectedIndex = 18 '교적카드
            SetMaximumOfferingNumberToolStripMenuItem.Visible = False
            cbCurrent.Visible = False
            cbPastor.Visible = True
            btCurrentGo1.Visible = False
            btGoPastor.Visible = True
        Else
            SetMaximumOfferingNumberToolStripMenuItem.Visible = True
            cbCurrent.SelectedIndex = 1 '주단위 헌금
            cbCurrent.Visible = True
            cbPastor.Visible = False
            btCurrentGo1.Visible = True
            btGoPastor.Visible = False
        End If

        tslblMode.Text = "현재 사용자: " + m_Global.loginUserName + "   "
        pnFind.Width = My.Computer.Screen.Bounds.Size.Width - gbCurrent.Width - 20

    End Sub
    Private Sub changeStatusCheckBox(ByVal flag1 As Boolean, ByVal flag2 As Boolean, ByVal flag3 As Boolean, ByVal flag4 As Boolean, ByVal flag5 As Boolean, ByVal flag6 As Boolean, ByVal flag7 As Boolean, ByVal flag8 As Boolean)
        chkStatus1.Checked = flag1
        chkStatus2.Checked = flag2
        chkStatus3.Checked = flag3
        chkStatus4.Checked = flag4

        chkStatus1.Enabled = flag5
        chkStatus2.Enabled = flag6
        chkStatus3.Enabled = flag7
        chkStatus4.Enabled = flag8
    End Sub

    Private Sub cbCurrent_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbCurrent.SelectedIndexChanged
        '0. 주단위 헌금 내역 (CRA)
        '1. 주단위 헌금 
        '2. 주단위 헌금 통계
        '3. 월단위 헌금 
        '4. 월단위 헌금 통계
        '5. 년단위 헌금 
        '6. 년단위 월별 헌금
        '7. 년단위 헌금 통계
        '8. 헌금영수증
        '9. 년도별 헌금봉투 정보 찾기
        '10. 부부 리스트(남편이 먼저)
        '11. 부부 리스트(부인이 먼저)
        '12. 가족 리스트
        '13. 고인 리스트
        '14. 유아 세례 리스트
        '15. 입교 리스트
        '16. 세례 리스트
        '17. 새교우 리스트
        '18. 등록 리스트
        '19. 이적 리스트
        '20. 생일 리스트
        '21. 교적카드
        '22. 헌금기록이 없는 교우 리스트
        '23. 교인 주소록 설문지

        If cbCurrent.SelectedIndex = 0 Or 'CRA
           cbCurrent.SelectedIndex = 1 Then '주단위 헌금
            cbCurrentEnableWhich(True, True)
            enableTypeCheckBox(cbCurrent.SelectedIndex)
            txtFind.SelectAll()
            txtFind.Focus()
            dtpCurrentStartDate.Visible = True
            dtpCurrentStartDate.Format = DateTimePickerFormat.Long
            dtpCurrentEndDate.Format = DateTimePickerFormat.Long
            changeStatusCheckBox(True, True, False, False, True, True, True, True)
            familyMode = False

        ElseIf cbCurrent.SelectedIndex = 2 Then '주단위 헌금통계
            cbCurrentEnableWhich(False, True)
            enableTypeCheckBox(cbCurrent.SelectedIndex)
            dtpCurrentStartDate.Visible = True
            dtpCurrentStartDate.Format = DateTimePickerFormat.Long
            dtpCurrentEndDate.Format = DateTimePickerFormat.Long
            changeStatusCheckBox(True, True, False, False, True, True, True, True)
            familyMode = False

        ElseIf cbCurrent.SelectedIndex = 3 Then '월단위 헌금
            cbCurrentEnableWhich(True, True)
            enableTypeCheckBox(cbCurrent.SelectedIndex)
            txtFind.SelectAll()
            txtFind.Focus()
            dtpCurrentStartDate.Visible = True
            dtpCurrentStartDate.Format = DateTimePickerFormat.Custom
            dtpCurrentStartDate.CustomFormat = "yyyy-MM"
            dtpCurrentEndDate.Format = DateTimePickerFormat.Custom
            dtpCurrentEndDate.CustomFormat = "yyyy-MM"
            changeStatusCheckBox(True, True, False, False, True, True, True, True)
            familyMode = False

        ElseIf cbCurrent.SelectedIndex = 4 Then '월단위 헌금통계
            cbCurrentEnableWhich(False, True)
            enableTypeCheckBox(cbCurrent.SelectedIndex)
            dtpCurrentStartDate.Visible = True
            dtpCurrentStartDate.Format = DateTimePickerFormat.Custom
            dtpCurrentStartDate.CustomFormat = "yyyy-MM"
            dtpCurrentEndDate.Format = DateTimePickerFormat.Custom
            dtpCurrentEndDate.CustomFormat = "yyyy-MM"
            changeStatusCheckBox(True, True, False, False, True, True, True, True)
            familyMode = False

        ElseIf cbCurrent.SelectedIndex = 5 Then '년단위 헌금
            cbCurrentEnableWhich(True, True)
            enableTypeCheckBox(cbCurrent.SelectedIndex)
            txtFind.SelectAll()
            txtFind.Focus()
            dtpCurrentStartDate.Visible = True
            dtpCurrentStartDate.Format = DateTimePickerFormat.Custom
            dtpCurrentStartDate.CustomFormat = "yyyy"
            dtpCurrentEndDate.Format = DateTimePickerFormat.Custom
            dtpCurrentEndDate.CustomFormat = "yyyy"
            changeStatusCheckBox(True, True, False, False, True, True, True, True)
            familyMode = False

        ElseIf cbCurrent.SelectedIndex = 6 Then  '년단위 월별 헌금
            cbCurrentEnableWhich(True, True)
            enableTypeCheckBox(cbCurrent.SelectedIndex)
            txtFind.SelectAll()
            txtFind.Focus()
            dtpCurrentStartDate.Visible = True
            dtpCurrentStartDate.Format = DateTimePickerFormat.Custom
            dtpCurrentStartDate.CustomFormat = "yyyy"
            dtpCurrentEndDate.Format = DateTimePickerFormat.Custom
            dtpCurrentEndDate.CustomFormat = "yyyy"
            changeStatusCheckBox(True, True, False, False, True, True, True, True)
            familyMode = False

        ElseIf cbCurrent.SelectedIndex = 7 Then '년단위 헌금통계
            cbCurrentEnableWhich(False, True)
            enableTypeCheckBox(cbCurrent.SelectedIndex)
            dtpCurrentStartDate.Visible = True
            dtpCurrentStartDate.Format = DateTimePickerFormat.Custom
            dtpCurrentStartDate.CustomFormat = "yyyy"
            dtpCurrentEndDate.Format = DateTimePickerFormat.Custom
            dtpCurrentEndDate.CustomFormat = "yyyy"
            changeStatusCheckBox(True, True, False, False, True, True, True, True)
            familyMode = False

        ElseIf cbCurrent.SelectedIndex = 8 Then '헌금영수증
            cbCurrentEnableWhich(True, True)
            enableTypeCheckBox(cbCurrent.SelectedIndex)
            txtFind.SelectAll()
            txtFind.Focus()
            dtpCurrentStartDate.Visible = False
            dtpCurrentStartDate.Format = DateTimePickerFormat.Custom
            dtpCurrentStartDate.CustomFormat = "yyyy"
            dtpCurrentEndDate.Format = DateTimePickerFormat.Custom
            dtpCurrentEndDate.CustomFormat = "yyyy"
            changeStatusCheckBox(True, True, False, False, True, True, True, True)
            familyMode = False

        ElseIf cbCurrent.SelectedIndex = 9 Then '특정년도 빈헌금봉투 리스트
            cbCurrentEnableWhich(True, True)
            enableTypeCheckBox(cbCurrent.SelectedIndex)
            dtpCurrentStartDate.Visible = False
            dtpCurrentEndDate.Format = DateTimePickerFormat.Custom
            dtpCurrentEndDate.CustomFormat = "yyyy"
            changeStatusCheckBox(True, True, False, False, True, True, True, True)
            familyMode = False

        ElseIf cbCurrent.SelectedIndex = 10 Or '남편이 먼저 부부리스트
               cbCurrent.SelectedIndex = 11 Or '부인이 먼저 부부리스트
               cbCurrent.SelectedIndex = 12 Then '가족 리스트
            dtpCurrentStartDate.Visible = True
            cbCurrentEnableWhich(True, False)
            changeStatusCheckBox(True, True, False, False, True, True, True, True)
            enableTypeCheckBox(cbCurrent.SelectedIndex)
            txtFind.SelectAll()
            txtFind.Focus()

            If cbCurrent.SelectedIndex = 10 Or cbCurrent.SelectedIndex = 11 Then
                familyMode = False
            End If

        ElseIf cbCurrent.SelectedIndex = 13 Or '사망자 리스트
            cbCurrent.SelectedIndex = 14 Or '유아세례 리스트
            cbCurrent.SelectedIndex = 15 Or '입교 리스트
            cbCurrent.SelectedIndex = 16 Or '세례 리스트
            cbCurrent.SelectedIndex = 17 Or ' 새교우 리스트
            cbCurrent.SelectedIndex = 18 Or '등록 리스트
            cbCurrent.SelectedIndex = 19 Then '이적 리스트
            dtpCurrentStartDate.Visible = True
            cbCurrentEnableWhich(True, True)
            enableTypeCheckBox(cbCurrent.SelectedIndex)
            txtFind.SelectAll()
            txtFind.Focus()
            dtpCurrentStartDate.Format = DateTimePickerFormat.Long
            dtpCurrentEndDate.Format = DateTimePickerFormat.Long

            If cbCurrent.SelectedIndex = 13 Then '고인
                changeStatusCheckBox(False, False, False, True, True, True, True, True)
            ElseIf cbCurrent.SelectedIndex = 17 Then '새교우
                changeStatusCheckBox(False, True, False, False, True, True, True, True)
            ElseIf cbCurrent.SelectedIndex = 19 Then '이적
                changeStatusCheckBox(False, False, True, False, True, True, True, True)
            ElseIf cbCurrent.SelectedIndex = 18 Then '등록
                changeStatusCheckBox(True, False, False, False, True, True, True, True)
            Else
                changeStatusCheckBox(True, True, False, False, True, True, True, True)
            End If

        ElseIf cbCurrent.SelectedIndex = 20 Then '생일리스트
            cbCurrentEnableWhich(True, True)
            enableTypeCheckBox(cbCurrent.SelectedIndex)
            txtFind.SelectAll()
            txtFind.Focus()
            dtpCurrentStartDate.Visible = True
            dtpCurrentStartDate.Format = DateTimePickerFormat.Custom
            dtpCurrentStartDate.CustomFormat = "M월"
            dtpCurrentEndDate.Format = DateTimePickerFormat.Custom
            dtpCurrentEndDate.CustomFormat = "M월"
            changeStatusCheckBox(True, True, False, False, True, True, True, True)

        ElseIf cbCurrent.SelectedIndex = 21 Then '교적카드
            cbCurrentEnableWhich(True, False)
            enableTypeCheckBox(cbCurrent.SelectedIndex)
            dtpCurrentStartDate.Visible = True
            txtFind.SelectAll()
            txtFind.Focus()
            changeStatusCheckBox(True, True, False, False, True, True, True, True)
            familyMode = True

        ElseIf cbCurrent.SelectedIndex = 22 Then '헌금기록이 없는 교우 리스트
            cbCurrentEnableWhich(True, True)
            enableTypeCheckBox(cbCurrent.SelectedIndex)
            dtpCurrentStartDate.Visible = True
            dtpCurrentStartDate.Format = DateTimePickerFormat.Long
            dtpCurrentEndDate.Format = DateTimePickerFormat.Long
            Dim d3weekagoSunday As Date = Now.AddDays(-(Now.DayOfWeek) - 21)
            dtpCurrentStartDate.Value = d3weekagoSunday
            changeStatusCheckBox(True, True, False, False, True, True, True, True)
            txtFind.SelectAll()
            txtFind.Focus()
            familyMode = False

        ElseIf cbCurrent.SelectedIndex = 23 Then '교인 주소록, 설문
            cbCurrentEnableWhich(True, False)
            enableTypeCheckBox(cbCurrent.SelectedIndex)
            dtpCurrentStartDate.Visible = True
            txtFind.SelectAll()
            txtFind.Focus()
            changeStatusCheckBox(True, True, False, False, True, True, True, True)

            'ElseIf cbCurrent.SelectedIndex = 25 Then '명찰
            '    cbCurrentEnableWhich(False, False)
            '    enableTypeCheckBox(cbCurrent.SelectedIndex)
            '    txtFind.SelectAll()
            '    txtFind.Focus()
            '    changeStatusCheckBox(True, True, False, False, False, False, False, False)
        End If
    End Sub

    Private Sub enableTypeCheckBox(ByVal index As Integer)
        '0. 주단위 헌금 내역 (CRA)
        '1. 주단위 헌금 
        '2. 주단위 헌금 통계
        '3. 월단위 헌금 
        '4. 월단위 헌금 통계
        '5. 년단위 헌금 
        '6. 년단위 월별 헌금
        '7. 년단위 헌금 통계
        '8. 헌금영수증
        '9. 년도별 헌금봉투 정보 찾기
        '10. 부부 리스트(남편이 먼저)
        '11. 부부 리스트(부인이 먼저)
        '12. 가족 리스트
        '13. 고인 리스트
        '14. 유아 세례 리스트
        '15. 입교 리스트
        '16. 세례 리스트
        '17. 새교우 리스트
        '18. 등록 리스트
        '19. 이적 리스트
        '20. 생일 리스트
        '21. 교적카드
        '22. 헌금기록이 없는 교우 리스트
        '23. 교인 주소록 설문지

        If index = 0 Then '0. 주단위 헌금 내역 (CRA)
            chkAll.Text = m_Constant.MODE_VIEW_OFFERING '"전체 개인별"
            chkType.Checked = True
            chkType.Enabled = False
            chkAll.Checked = False
            chkAll.Enabled = False
        ElseIf index = 1 Or index = 3 Or index = 5 Then '1. 단위 헌금 
            chkAll.Text = m_Constant.MODE_VIEW_OFFERING '"전체 개인별"
            chkType.Checked = True
            chkType.Enabled = True
            chkAll.Enabled = True
        ElseIf index = 2 Or index = 4 Or index = 7 Then '2. 통계
            chkAll.Text = m_Constant.MODE_VIEW_OFFERING '"전체 개인별"
            chkType.Checked = False
            chkType.Enabled = False
            chkAll.Checked = False
            chkAll.Enabled = False
        ElseIf index = 6 Then '6. 년단위 월별 헌금
            chkAll.Text = m_Constant.MODE_VIEW_OFFERING '"전체 개인별"
            chkType.Checked = True
            chkType.Enabled = True
            chkAll.Checked = False
            chkAll.Enabled = False
        ElseIf index = 8 Then '헌금 영수증
            chkAll.Text = m_Constant.MODE_VIEW_OFFERING '"전체 개인별"
            chkType.Checked = False
            chkType.Enabled = False
            chkAll.Checked = True
            chkAll.Enabled = True
        ElseIf index = 9 Then '특정년도 빈헌금봉투 리스트
            chkAll.Text = m_Constant.MODE_VIEW_OFFERING '"전체 개인별"
            chkType.Checked = False
            chkType.Enabled = False
            chkAll.Checked = True
            chkAll.Enabled = True
        ElseIf index = 10 Or index = 11 Then '부부 리스트
            chkAll.Text = m_Constant.MODE_VIEW_OFFERING '"전체 개인별"
            chkType.Checked = False
            chkType.Enabled = False
            chkAll.Checked = True
            chkAll.Enabled = True
        ElseIf index = 12 Then '가족 리스트
            chkAll.Text = m_Constant.MODE_VIEW_FAMILY '"가족 모드로 보기"
            chkType.Checked = False
            chkType.Enabled = False
            chkAll.Checked = True
            chkAll.Enabled = False
        ElseIf index = 13 Or
            index = 14 Or index = 15 Or index = 16 Or index = 17 Or
            index = 18 Or index = 19 Or index = 20 Then
            chkAll.Text = m_Constant.MODE_VIEW_FAMILY '"가족 모드로 보기"
            chkType.Checked = False
            chkType.Enabled = False
            chkAll.Checked = False
            chkAll.Enabled = True
        ElseIf index = 21 Or index = 23 Then
            chkAll.Text = m_Constant.MODE_VIEW_FAMILY '"가족 모드로 보기"
            chkType.Checked = False
            chkType.Enabled = False
            chkAll.Checked = True
            chkAll.Enabled = False
        ElseIf index = 22 Then
            chkAll.Text = m_Constant.MODE_VIEW_OFFERING '"개인 전체별"
            chkType.Checked = False
            chkType.Enabled = False
            chkAll.Checked = False
            chkAll.Enabled = False
        End If
    End Sub

    Private Sub enableTypeCheckBoxPastor(ByVal index As Integer)
        '0-주단위 헌금 집계 현황
        '1-주단위 헌금 통계
        '2-월단위 헌금 집계 현황
        '3-월단위 헌금 통계
        '4-년단위 헌금 집계 현황
        '5-년단위 월별 헌금
        '6-년단위 헌금 통계
        '7-부부 리스트(남편이 먼저)
        '8-부부 리스트(부인이 먼저)
        '9-가족 리스트
        '10-고인 리스트
        '11-유아세례 리스트
        '12-입교 리스트
        '13-세례 리스트
        '14-새교우 리스트
        '15-등록교인 리스트
        '16-이적교인 리스트
        '17-생일 리스트
        '18-교적카드
        '19-헌금기록이 없는 교우 리스트
        '20-교인 주소록 설문 양식

        If index = 0 Or index = 2 Or index = 4 Then '1. 단위 헌금 
            chkAll.Text = "전체 개인별"
            chkType.Checked = True
            chkType.Enabled = True
            chkAll.Checked = False
            chkAll.Enabled = False
        ElseIf index = 1 Or index = 3 Or index = 6 Then '2. 통계
            chkAll.Text = "전체 개인별"
            chkType.Checked = False
            chkType.Enabled = False
            chkAll.Checked = False
            chkAll.Enabled = False
        ElseIf index = 5 Then '6. 년단위 월별 헌금
            chkAll.Text = "전체 개인별"
            chkType.Checked = True
            chkType.Enabled = True
            chkAll.Checked = False
            chkAll.Enabled = False
        ElseIf index = 7 Or index = 8 Then '부부 리스트
            chkAll.Text = "전체 개인별"
            chkType.Checked = False
            chkType.Enabled = False
            chkAll.Checked = True
            chkAll.Enabled = True
        ElseIf index = 9 Then '가족 리스트
            chkAll.Text = "가족 모드로 보기"
            chkType.Checked = False
            chkType.Enabled = False
            chkAll.Checked = True
            chkAll.Enabled = False
        ElseIf index = 10 Or
            index = 11 Or index = 12 Or index = 13 Or index = 14 Or
            index = 15 Or index = 16 Or index = 17 Then
            chkAll.Text = "가족 모드로 보기"
            chkType.Checked = False
            chkType.Enabled = False
            chkAll.Checked = False
            chkAll.Enabled = True
        ElseIf index = 18 Or index = 20 Then
            chkAll.Text = "가족 모드로 보기"
            chkType.Checked = False
            chkType.Enabled = False
            chkAll.Checked = True
            chkAll.Enabled = False
        End If
    End Sub

    Private Sub cbCurrentEnableWhich(ByRef flag1 As Boolean, ByRef flag2 As Boolean)
        pnCurrent1.Enabled = flag1
        pnCurrent2.Enabled = flag2
    End Sub

    Private Sub initDgChosen()
        With dgChosen
            .Columns(0).Width = 60 'Person id
            .Columns(0).Visible = False
            .Columns(1).Width = 80 'Korean name
            .Columns(2).Width = 60 'Spouses
            .Columns(2).Visible = False
            .Columns(3).Width = 60
            .Columns(4).Width = 60
            .Font = chkAll.Font
        End With
    End Sub

    Private Sub dgChosen_KeyUp(sender As Object, e As KeyEventArgs) Handles dgChosen.KeyUp
        If e.KeyCode = Keys.Delete Then
            If dgChosen.RowCount = 1 Then Exit Sub
            Dim result As Integer = MessageBox.Show("선택한 레코드를 삭제하시겠습니까?", "삭제", MessageBoxButtons.YesNo)
            If result = DialogResult.Yes Then
                Dim personId As String = CType(dgChosen.Rows(dgChosen.CurrentRow.Index).Cells(0).Value, String)
                If findPeopleHashTable.ContainsKey(personId) Then
                    findPeopleHashTable.Remove(personId)
                End If
                dgChosen.Rows.RemoveAt(dgChosen.CurrentRow.Index)
            End If
        End If
    End Sub

    Private Sub btCurrentGo1_Click(sender As Object, e As EventArgs) Handles btCurrentGo1.Click
        txtFind.Clear()
        pnFind.Visible = False

        Dim title As String = ""

        If cbCurrent.SelectedIndex = 0 Or cbCurrent.SelectedIndex = 1 Or
           cbCurrent.SelectedIndex = 2 Or cbCurrent.SelectedIndex = 3 Or
           cbCurrent.SelectedIndex = 4 Or cbCurrent.SelectedIndex = 5 Or
           cbCurrent.SelectedIndex = 6 Or cbCurrent.SelectedIndex = 7 Then '주단위, 년단위
            Dim offeringDAL As TKPC.DAL.OfferingDAL = TKPC.DAL.OfferingDAL.getInstance()
            Dim ds As DataSet = Nothing
            Dim selectedStartDate As String = dtpCurrentStartDate.Value.ToString("yyyy-MM-dd")
            Dim selectedEndDate As String = dtpCurrentEndDate.Value.ToString("yyyy-MM-dd")

            Try
                If m_Global.compare2Dates(CType(selectedStartDate, Date), CType(selectedEndDate, Date)) = False Then
                    MsgBox("시작날짜가 끝날짜보다 더 클 수 없습니다. 다시 선택해주십시오.", Nothing, "주의")
                    dtpCurrentStartDate.Focus()
                    Exit Sub
                End If

                Me.Cursor = Cursors.WaitCursor

                Dim personIds As String = ""
                Dim cp As List(Of String) = readPersonIdFromDgChosen(0)
                If cp.Count > 0 Then
                    personIds = String.Join(",", cp)
                    prevFoundPersonList = cp
                End If

                If cbCurrent.SelectedIndex = 0 Then
                    If String.IsNullOrEmpty(personIds) = True Then
                        MsgBox("CRA보고서를 위해 교우 한사람을 선택해주시기 바랍니다.", Nothing, "주의")
                        txtFind.Focus()
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    ElseIf dgChosen.RowCount > 1 Then
                        MsgBox("교우 한사람만 선택해주시기 바랍니다.", Nothing, "주의")
                        txtFind.SelectAll()
                        txtFind.Focus()
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    Else
                        ds = offeringDAL.getOfferingFlexibleReport(True, selectedStartDate, selectedEndDate, personIds)
                    End If

                ElseIf cbCurrent.SelectedIndex = 1 Then '주단위 헌금 집계 현황
                    If chkType.Checked = False And chkAll.Checked = False And dgChosen.RowCount = 0 Then
                        chkType.Checked = True 'default
                    End If

                    Dim datesList As List(Of String) = m_Global.getOfferingDates(0, selectedStartDate, selectedEndDate)
                    ds = offeringDAL.getOfferingFlexibleReport1(chkType.Checked, selectedStartDate, selectedEndDate, personIds, chkAll.Checked, datesList)

                ElseIf cbCurrent.SelectedIndex = 2 Then '주단위 헌금통계
                    ds = offeringDAL.getOfferingFlexibleReport2(selectedStartDate, selectedEndDate)

                ElseIf cbCurrent.SelectedIndex = 3 Then '월단위 헌금 집계 현황
                    If chkType.Checked = False And chkAll.Checked = False And dgChosen.RowCount = 0 Then
                        chkType.Checked = True 'default
                    End If

                    selectedStartDate = m_Global.formatDate(1, selectedStartDate, "01")
                    selectedEndDate = m_Global.formatDate(1, selectedEndDate, "31")

                    ds = offeringDAL.getOfferingFlexibleReport3(chkType.Checked, selectedStartDate, selectedEndDate, personIds, chkAll.Checked)

                ElseIf cbCurrent.SelectedIndex = 4 Then '월단위 헌금통계
                    selectedStartDate = m_Global.formatDate(1, selectedStartDate, "01")
                    selectedEndDate = m_Global.formatDate(1, selectedEndDate, "31")

                    ds = offeringDAL.getOfferingFlexibleReport4(selectedStartDate, selectedEndDate)

                ElseIf cbCurrent.SelectedIndex = 5 Then '년단위 헌금 집계
                    If chkType.Checked = False And chkAll.Checked = False And dgChosen.RowCount = 0 Then
                        chkType.Checked = True 'default
                    End If
                    selectedStartDate = dtpCurrentStartDate.Value.ToString("yyyy")
                    selectedEndDate = dtpCurrentEndDate.Value.ToString("yyyy")

                    ds = offeringDAL.getOfferingFlexibleReport5(chkType.Checked, selectedStartDate, selectedEndDate, personIds, chkAll.Checked)

                ElseIf cbCurrent.SelectedIndex = 6 Then '년단위 월별 헌금 통계
                    selectedStartDate = dtpCurrentStartDate.Value.ToString("yyyy")
                    selectedEndDate = dtpCurrentEndDate.Value.ToString("yyyy")
                    ds = offeringDAL.getOfferingFlexibleReport6(chkType.Checked, selectedStartDate, selectedEndDate, personIds)

                ElseIf cbCurrent.SelectedIndex = 7 Then '년단위 헌금 통계
                    selectedStartDate = dtpCurrentStartDate.Value.ToString("yyyy")
                    selectedEndDate = dtpCurrentEndDate.Value.ToString("yyyy")
                    ds = offeringDAL.getOfferingFlexibleReport7(selectedStartDate, selectedEndDate)
                End If

                title = cbCurrent.Text + ", 기간: " + selectedStartDate + " ~ " + selectedEndDate

                m_Global.setNavigatorBinding(ds, 1, False, title)

            Catch ex As Exception
                Logger.LogError(ex.ToString)
            Finally
                Me.Cursor = Cursors.Default
                offeringDAL = Nothing
                ds = Nothing
            End Try

        ElseIf cbCurrent.SelectedIndex = 8 Then '헌금영수증
            tsbtCard.Visible = False
            tsbtSurvey.Visible = False
            frmReceiptInfo.flag = 0
            frmReceiptInfo.ShowDialog()

        ElseIf cbCurrent.SelectedIndex = 9 Then '년도별 헌금봉투 찾기
            Dim endYear As String = dtpCurrentEndDate.Value.ToString("yyyy")

            Me.Cursor = Cursors.WaitCursor
            Dim offeringDAL As TKPC.DAL.OfferingDAL = TKPC.DAL.OfferingDAL.getInstance()
            Dim ds As DataSet = Nothing

            Try
                If cbFindWhat.SelectedIndex = 0 Then
                    Dim personIds As String = ""
                    Dim cp As List(Of String) = readPersonIdFromDgChosen(0)
                    If cp.Count > 0 Then
                        personIds = String.Join(",", cp)
                        prevFoundPersonList = cp
                    End If
                    ds = offeringDAL.getOfferingInfoByPersonIdOrOfferingId(cbFindWhat.SelectedIndex, personIds, "", endYear)
                Else
                    ds = offeringDAL.getOfferingInfoByPersonIdOrOfferingId(cbFindWhat.SelectedIndex, Nothing, txtFind.Text, endYear)
                End If

                title = cbCurrent.Text + ", 년도: " + endYear
                m_Global.setNavigatorBinding(ds, 1, False, title)

            Catch ex As Exception
                Logger.LogError(ex.ToString)
            Finally
                offeringDAL = Nothing
                ds = Nothing
                Me.Cursor = Cursors.Default
            End Try

        ElseIf cbCurrent.SelectedIndex = 10 Or cbCurrent.SelectedIndex = 11 Then '부부리스트
            Me.Cursor = Cursors.WaitCursor
            Dim personDAL As TKPC.DAL.PersonDAL = TKPC.DAL.PersonDAL.getInstance()
            Dim ds As DataSet = Nothing

            Try
                Dim personNames As String = ""
                Dim cp As List(Of String) = readPersonIdFromDgChosen(1)
                If cp.Count > 0 Then
                    personNames = String.Join("','", cp)
                    personNames = "'" + personNames + "'"
                    prevFoundPersonList = cp
                End If

                Dim statusOptions As String = ""
                '0 - 전체 데이터
                If chkStatus1.Checked = True And chkStatus2.Checked = True And chkStatus3.Checked = True And chkStatus4.Checked = True Then
                    ds = personDAL.getCoupleList(cbCurrent.SelectedIndex - 10, personNames, 0)
                    statusOptions = "전체"

                    '1 - 등록교인만
                ElseIf chkStatus1.Checked = True And chkStatus2.Checked = False And chkStatus3.Checked = False And chkStatus4.Checked = False Then
                    ds = personDAL.getCoupleList(cbCurrent.SelectedIndex - 10, personNames, 1)
                    statusOptions = "등록교인"

                    '2 - 등록교인 + 새교우
                ElseIf chkStatus1.Checked = True And chkStatus2.Checked = True And chkStatus3.Checked = False And chkStatus4.Checked = False Then
                    ds = personDAL.getCoupleList(cbCurrent.SelectedIndex - 10, personNames, 2)
                    statusOptions = "등록교인,새교우"

                    '3 - 등록교인 + 새교우 + 이적교인
                ElseIf chkStatus1.Checked = True And chkStatus2.Checked = True And chkStatus3.Checked = True And chkStatus4.Checked = False Then
                    ds = personDAL.getCoupleList(cbCurrent.SelectedIndex - 10, personNames, 3)
                    statusOptions = "등록교인,새교우,이적교인"

                    '4 - 등록교인 + 새교우 + 고인
                ElseIf chkStatus1.Checked = True And chkStatus2.Checked = True And chkStatus3.Checked = False And chkStatus4.Checked = True Then
                    ds = personDAL.getCoupleList(cbCurrent.SelectedIndex - 10, personNames, 4)
                    statusOptions = "등록교인,새교우,고인"

                    '5 - 등록교인 + 이적교인
                ElseIf chkStatus1.Checked = True And chkStatus2.Checked = False And chkStatus3.Checked = True And chkStatus4.Checked = False Then
                    ds = personDAL.getCoupleList(cbCurrent.SelectedIndex - 10, personNames, 5)
                    statusOptions = "등록교인,이적교인"

                    '6 - 등록교인 + 이적교인 + 고인
                ElseIf chkStatus1.Checked = True And chkStatus2.Checked = False And chkStatus3.Checked = True And chkStatus4.Checked = True Then
                    ds = personDAL.getCoupleList(cbCurrent.SelectedIndex - 10, personNames, 6)
                    statusOptions = "등록교인,이적교인,고인"

                    '7 - 등록교인 + 고인
                ElseIf chkStatus1.Checked = True And chkStatus2.Checked = False And chkStatus3.Checked = False And chkStatus4.Checked = True Then
                    ds = personDAL.getCoupleList(cbCurrent.SelectedIndex - 10, personNames, 7)
                    statusOptions = "등록교인,고인"

                    '8 - 새교우만
                ElseIf chkStatus1.Checked = False And chkStatus2.Checked = True And chkStatus3.Checked = False And chkStatus4.Checked = False Then
                    ds = personDAL.getCoupleList(cbCurrent.SelectedIndex - 10, personNames, 8)
                    statusOptions = "새교우"

                    '9 - 새교우 + 이적교인
                ElseIf chkStatus1.Checked = False And chkStatus2.Checked = True And chkStatus3.Checked = True And chkStatus4.Checked = False Then
                    ds = personDAL.getCoupleList(cbCurrent.SelectedIndex - 10, personNames, 9)
                    statusOptions = "새교우,이적교인"

                    '10 - 새교우 + 이적교인 + 고인
                ElseIf chkStatus1.Checked = False And chkStatus2.Checked = True And chkStatus3.Checked = True And chkStatus4.Checked = True Then
                    ds = personDAL.getCoupleList(cbCurrent.SelectedIndex - 10, personNames, 10)
                    statusOptions = "새교우,이적교인,고인"

                    '11 - 새교우 + 고인
                ElseIf chkStatus1.Checked = False And chkStatus2.Checked = True And chkStatus3.Checked = False And chkStatus4.Checked = True Then
                    ds = personDAL.getCoupleList(cbCurrent.SelectedIndex - 10, personNames, 11)
                    statusOptions = "새교우,고인"

                    '12 - 이적교인만
                ElseIf chkStatus1.Checked = False And chkStatus2.Checked = False And chkStatus3.Checked = True And chkStatus4.Checked = False Then
                    ds = personDAL.getCoupleList(cbCurrent.SelectedIndex - 10, personNames, 12)
                    statusOptions = "이적교인"

                    '13 - 이적교인 + 고인
                ElseIf chkStatus1.Checked = False And chkStatus2.Checked = False And chkStatus3.Checked = True And chkStatus4.Checked = True Then
                    ds = personDAL.getCoupleList(cbCurrent.SelectedIndex - 10, personNames, 13)
                    statusOptions = "이적교인,고인"

                    '14 - 고인만
                ElseIf chkStatus1.Checked = False And chkStatus2.Checked = False And chkStatus3.Checked = False And chkStatus4.Checked = True Then
                    ds = personDAL.getCoupleList(cbCurrent.SelectedIndex - 10, personNames, 14)
                    statusOptions = "고인"

                    '15 - 모두 아님
                ElseIf chkStatus1.Checked = False And chkStatus2.Checked = False And chkStatus3.Checked = False And chkStatus4.Checked = False Then
                    chkStatus1.Checked = True
                    ds = personDAL.getCoupleList(cbCurrent.SelectedIndex - 10, personNames, 15)
                    statusOptions = "등록교인"
                End If

                title = cbCurrent.Text + ", 등록옵션: " + statusOptions

                m_Global.setNavigatorBinding(ds, 1, False, title)
            Catch ex As Exception
                Logger.LogError(ex.ToString)
            Finally
                personDAL = Nothing
                ds = Nothing
                Me.Cursor = Cursors.Default
            End Try

        ElseIf cbCurrent.SelectedIndex = 12 Then '가족 리스트
            familyMode = True
            prevFoundPersonList = readPersonIdFromDgChosen(0)
            getFamilyList(0, readPersonIdFromDgChosen(0))
            familyMode = False

        ElseIf cbCurrent.SelectedIndex = 13 Or   '고인 리스트
               cbCurrent.SelectedIndex = 14 Or   '유아세례 리스트
               cbCurrent.SelectedIndex = 15 Or   '입교 리스트
               cbCurrent.SelectedIndex = 16 Or   '세례 리스트
               cbCurrent.SelectedIndex = 17 Or   '새교우 리스트
               cbCurrent.SelectedIndex = 18 Or   '등록 리스트
               cbCurrent.SelectedIndex = 19 Or   '이적 리스트
               cbCurrent.SelectedIndex = 20 Then '생일 리스트

            Dim personDAL As TKPC.DAL.PersonDAL = TKPC.DAL.PersonDAL.getInstance()
            Dim ds As DataSet = Nothing
            Dim startDate As String = dtpCurrentStartDate.Text
            Dim endDate As String = dtpCurrentEndDate.Text

            If cbCurrent.SelectedIndex = 20 Then
                Dim startMonth As Integer = CType(startDate.Substring(0, startDate.Length - 1), Integer)
                Dim endMonth As Integer = CType(endDate.Substring(0, endDate.Length - 1), Integer)

                startDate = CType(startDate.Substring(0, startDate.Length - 1), Integer).ToString("D2")
                endDate = CType(endDate.Substring(0, endDate.Length - 1), Integer).ToString("D2")

                If m_Global.compare(startMonth, endMonth) = False Then
                    MsgBox("시작날짜가 끝날짜보다 더 클 수 없습니다. 다시 선택해주십시오.", Nothing, "주의")
                    dtpCurrentStartDate.Focus()
                    Exit Sub
                End If
            Else
                startDate = dtpCurrentStartDate.Value.ToString("yyyy-MM-dd")
                endDate = dtpCurrentEndDate.Value.ToString("yyyy-MM-dd")

                If m_Global.compare2Dates(CType(startDate, Date), CType(endDate, Date)) = False Then
                    MsgBox("시작날짜가 끝날짜보다 더 클 수 없습니다. 다시 선택해주십시오.", Nothing, "주의")
                    dtpCurrentStartDate.Focus()
                    Exit Sub
                End If
            End If

            Try
                Me.Cursor = Cursors.WaitCursor

                Dim personIds As String = ""
                Dim cp As List(Of String) = readPersonIdFromDgChosen(0)
                If cp.Count > 0 Then
                    personIds = String.Join(",", cp)
                    prevFoundPersonList = cp
                End If

                Dim statusOptions As String = ""
                '0 - 전체 데이터
                If chkStatus1.Checked = True And chkStatus2.Checked = True And chkStatus3.Checked = True And chkStatus4.Checked = True Then
                    ds = personDAL.getPeopleList(cbCurrent.SelectedIndex - 13, personIds, startDate, endDate, 0)
                    statusOptions = "전체"

                    '1 - 등록교인만
                ElseIf chkStatus1.Checked = True And chkStatus2.Checked = False And chkStatus3.Checked = False And chkStatus4.Checked = False Then
                    ds = personDAL.getPeopleList(cbCurrent.SelectedIndex - 13, personIds, startDate, endDate, 1)
                    statusOptions = "등록교인"

                    '2 - 등록교인 + 새교우
                ElseIf chkStatus1.Checked = True And chkStatus2.Checked = True And chkStatus3.Checked = False And chkStatus4.Checked = False Then
                    ds = personDAL.getPeopleList(cbCurrent.SelectedIndex - 13, personIds, startDate, endDate, 2)
                    statusOptions = "등록교인,새교우"

                    '3 - 등록교인 + 새교우 + 이적교인
                ElseIf chkStatus1.Checked = True And chkStatus2.Checked = True And chkStatus3.Checked = True And chkStatus4.Checked = False Then
                    ds = personDAL.getPeopleList(cbCurrent.SelectedIndex - 13, personIds, startDate, endDate, 3)
                    statusOptions = "등록교인,새교우,이적교인"

                    '4 - 등록교인 + 새교우 + 고인
                ElseIf chkStatus1.Checked = True And chkStatus2.Checked = True And chkStatus3.Checked = False And chkStatus4.Checked = True Then
                    ds = personDAL.getPeopleList(cbCurrent.SelectedIndex - 13, personIds, startDate, endDate, 4)
                    statusOptions = "등록교인,새교우,고인"

                    '5 - 등록교인 + 이적교인
                ElseIf chkStatus1.Checked = True And chkStatus2.Checked = False And chkStatus3.Checked = True And chkStatus4.Checked = False Then
                    ds = personDAL.getPeopleList(cbCurrent.SelectedIndex - 13, personIds, startDate, endDate, 5)
                    statusOptions = "등록교인,이적교인"

                    '6 - 등록교인 + 이적교인 + 고인
                ElseIf chkStatus1.Checked = True And chkStatus2.Checked = False And chkStatus3.Checked = True And chkStatus4.Checked = True Then
                    ds = personDAL.getPeopleList(cbCurrent.SelectedIndex - 13, personIds, startDate, endDate, 6)
                    statusOptions = "등록교인,이적교인,고인"

                    '7 - 등록교인 + 고인
                ElseIf chkStatus1.Checked = True And chkStatus2.Checked = False And chkStatus3.Checked = False And chkStatus4.Checked = True Then
                    ds = personDAL.getPeopleList(cbCurrent.SelectedIndex - 13, personIds, startDate, endDate, 7)
                    statusOptions = "등록교인,고인"

                    '8 - 새교우만
                ElseIf chkStatus1.Checked = False And chkStatus2.Checked = True And chkStatus3.Checked = False And chkStatus4.Checked = False Then
                    ds = personDAL.getPeopleList(cbCurrent.SelectedIndex - 13, personIds, startDate, endDate, 8)
                    statusOptions = "새교우"

                    '9 - 새교우 + 이적교인
                ElseIf chkStatus1.Checked = False And chkStatus2.Checked = True And chkStatus3.Checked = True And chkStatus4.Checked = False Then
                    ds = personDAL.getPeopleList(cbCurrent.SelectedIndex - 13, personIds, startDate, endDate, 9)
                    statusOptions = "새교우,이적교인"

                    '10 - 새교우 + 이적교인 + 고인
                ElseIf chkStatus1.Checked = False And chkStatus2.Checked = True And chkStatus3.Checked = True And chkStatus4.Checked = True Then
                    ds = personDAL.getPeopleList(cbCurrent.SelectedIndex - 13, personIds, startDate, endDate, 10)
                    statusOptions = "새교우,이적교인,고인"

                    '11 - 새교우 + 고인
                ElseIf chkStatus1.Checked = False And chkStatus2.Checked = True And chkStatus3.Checked = False And chkStatus4.Checked = True Then
                    ds = personDAL.getPeopleList(cbCurrent.SelectedIndex - 13, personIds, startDate, endDate, 11)
                    statusOptions = "새교우,고인"

                    '12 - 이적교인만
                ElseIf chkStatus1.Checked = False And chkStatus2.Checked = False And chkStatus3.Checked = True And chkStatus4.Checked = False Then
                    ds = personDAL.getPeopleList(cbCurrent.SelectedIndex - 13, personIds, startDate, endDate, 12)
                    statusOptions = "이적교인"

                    '13 - 이적교인 + 고인
                ElseIf chkStatus1.Checked = False And chkStatus2.Checked = False And chkStatus3.Checked = True And chkStatus4.Checked = True Then
                    ds = personDAL.getPeopleList(cbCurrent.SelectedIndex - 13, personIds, startDate, endDate, 13)
                    statusOptions = "이적교인,고인"

                    '14 - 고인만
                ElseIf chkStatus1.Checked = False And chkStatus2.Checked = False And chkStatus3.Checked = False And chkStatus4.Checked = True Then
                    ds = personDAL.getPeopleList(cbCurrent.SelectedIndex - 13, personIds, startDate, endDate, 14)
                    statusOptions = "고인"

                    '15 - 모두 아님
                ElseIf chkStatus1.Checked = False And chkStatus2.Checked = False And chkStatus3.Checked = False And chkStatus4.Checked = False Then
                    chkStatus1.Checked = True
                    ds = personDAL.getPeopleList(cbCurrent.SelectedIndex - 13, personIds, startDate, endDate, 15)
                    statusOptions = "등록교인"

                End If

                title = cbCurrent.Text + ", 등록옵션: " + statusOptions + ", 기간: " + startDate + " ~ " + endDate
                m_Global.setNavigatorBinding(ds, 1, True, title)

                If chkAll.Checked = True Then '가족모드로 보기
                    familyMode = True
                    dgChosen.Rows.Clear()

                    If ugResult.Rows.Count > 0 Then
                        With ugResult
                            For i As Integer = 0 To .Rows.Count - 1

                                Dim status As String = .Rows(i).Cells("status").Text
                                Dim deceased As String = .Rows(i).Cells("deceased_date").Text
                                If String.IsNullOrEmpty(deceased) = True Then
                                    If status = "church_member" Then
                                        status = "등록"
                                    ElseIf status = "left_church" Then
                                        status = "이적"
                                    ElseIf status = "new_family" Then
                                        status = "새교우"
                                    End If
                                Else
                                    status = "고인"
                                End If

                                dgChosen.Rows.Add(CType(.Rows(i).Cells("person_id").Value, String),
                                                  .Rows(i).Cells("korean_name").Text,
                                              CType(.Rows(i).Cells("spouses").Value, String),
                                              status,
                                              .Rows(i).Cells("age").Text,
                                              .Rows(i).Cells("gender").Text)
                            Next
                        End With

                        getFamilyList(0, readPersonIdFromDgChosen(0))

                    End If
                Else
                    tsbtSurvey.Visible = False
                    tsbtCard.Visible = False
                    familyMode = False
                End If

            Catch ex As Exception
                MessageBox.Show(ex.ToString)
                Logger.LogError(ex.ToString)
            Finally
                personDAL = Nothing
                ds = Nothing
                Me.Cursor = Cursors.Default
            End Try

        ElseIf cbCurrent.SelectedIndex = 21 Then '교적카드
            'If dgChosen.RowCount = 0 Then
            '    If MsgBox("데이터의 양이 많아?", MsgBoxStyle.YesNo, "백그라운드 처리") = MsgBoxResult.Yes Then ' If you select yes in the MsgBox then it will close the window
            '        bwTKPC.RunWorkerAsync(3000)
            '    End If
            'Else
            If dgChosen.Rows.Count = 0 Then
                MsgBox("교우를 선택해주시기 바랍니다.", Nothing, "주의")
                txtFind.Focus()
            Else
                familyMode = True
                writeMembershipCard()
                familyMode = False
            End If

            'End If
        ElseIf cbCurrent.SelectedIndex = 22 Then '헌금기록이 없는 사람들
            Me.Cursor = Cursors.WaitCursor
            Dim offeringDAL As TKPC.DAL.OfferingDAL = TKPC.DAL.OfferingDAL.getInstance()
            Dim dt As DataTable = Nothing

            Try
                Dim startDate As String = dtpCurrentStartDate.Value.ToString("yyyy-MM-dd")
                Dim endDate As String = dtpCurrentEndDate.Value.ToString("yyyy-MM-dd")

                If m_Global.compare2Dates(CType(startDate, Date), CType(endDate, Date)) = False Then
                    MsgBox("시작날짜가 끝날짜보다 더 클 수 없습니다. 다시 선택해주십시오.", Nothing, "주의")
                    dtpCurrentStartDate.Focus()
                    Me.Cursor = Cursors.Default
                    Exit Sub
                End If

                Dim personIds As String = ""
                Dim cp As List(Of String) = readPersonIdFromDgChosen(0)
                If cp.Count > 0 Then
                    personIds = String.Join(",", cp)
                End If

                Dim startYear As Integer = CType(dtpCurrentStartDate.Value.ToString("yyyy"), Integer)
                Dim endYear As Integer = CType(dtpCurrentEndDate.Value.ToString("yyyy"), Integer)
                Dim paramYears As String = ""

                For i As Integer = startYear To endYear
                    If i < endYear Then
                        paramYears += CType(i, String) + ","
                    Else
                        paramYears += CType(i, String)
                    End If
                Next

                dt = offeringDAL.getPeopleWithoutOffering(startDate, endDate, personIds, paramYears)
                With ugResult
                    .DataSource = dt
                    .DataBind()
                    .Text = cbCurrent.Text + ", 기간: " + startDate + " ~ " + endDate
                End With

            Catch ex As Exception
                Logger.LogError(ex.ToString)
            Finally
                offeringDAL = Nothing
                dt = Nothing
                Me.Cursor = Cursors.Default
            End Try

            'ElseIf cbCurrent.SelectedIndex = 23 Then '교인 주소록
            '    MessageBox.Show("Coming!")
            '    'writeDirectoryBook()
        ElseIf cbCurrent.SelectedIndex = 23 Then '교인 주소록 설문
            familyMode = True
            writeSurveyForm(0, -1)
            familyMode = False
            'ElseIf cbCurrent.SelectedIndex = 25 Then '명찰
            '    MessageBox.Show("Coming!")
            '    'frmNameTagOption.ShowDialog()
        End If
    End Sub

    Private Sub dgChosen_RowsRemoved(sender As Object, e As DataGridViewRowsRemovedEventArgs) Handles dgChosen.RowsRemoved
        If dgChosen.RowCount > 0 Then
            enableUpdateAll(True, True, True, True, True)
            If m_Global.accessLevel = m_Constant.ACCESS_LEVEL_OFFICE Then
                chkAll.Checked = False
            End If
        Else
            enableUpdateAll(False, False, False, False, False)
            If m_Global.accessLevel = m_Constant.ACCESS_LEVEL_OFFICE Then
                chkAll.Checked = True
            End If
        End If
    End Sub

    Private Sub dgChosen_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs) Handles dgChosen.RowsAdded
        If dgChosen.RowCount > 0 Then
            enableUpdateAll(True, True, True, True, True)

            If m_Global.accessLevel = m_Constant.ACCESS_LEVEL_OFFICE Then
                If cbCurrent.SelectedIndex = 1 Then
                    chkAll.Checked = False
                End If
            End If
        Else
            enableUpdateAll(False, False, False, False, False)
        End If
    End Sub

    Private Sub writeMembershipCard()
        Dim cp As List(Of String) = readPersonIdFromDgChosen(0)
        If ugResult.Rows.Count = 0 Then
            m_Global.getFamilyList(0, cp)
        Else
            Dim str1 As String = ""
            Dim str2 As String = ""
            If cp.Count > 0 Then
                str1 = String.Join(",", cp)
            End If

            If prevFoundPersonList.Count > 0 Then
                str2 = String.Join(",", prevFoundPersonList)
            End If

            If ugResult.DisplayLayout.Bands(0).Columns("제외").Header.Caption = "제외" Then
                If str1 <> str2 Then
                    m_Global.getFamilyList(0, cp)
                End If
            Else
                m_Global.getFamilyList(0, cp)
            End If
        End If

        m_Global.writeMembershipCard(0, Nothing, False)
        prevFoundPersonList = cp
    End Sub

    Private Sub writeSurveyForm(ByVal htmlFlag As Integer, ByVal labelOption As Integer)
        Dim cp As List(Of String) = readPersonIdFromDgChosen(0)
        If ugResult.Rows.Count = 0 And htmlFlag = 0 Then
            m_Global.getFamilyList(0, cp)
        Else
            Dim str1 As String = ""
            Dim str2 As String = ""
            If cp.Count > 0 Then
                str1 = String.Join(",", cp)
            End If

            If prevFoundPersonList.Count > 0 Then
                str2 = String.Join(",", prevFoundPersonList)
            End If

            If ugResult.DisplayLayout.Bands(0).Columns("제외").Header.Caption = "제외" Then
                If str1 <> str2 Then
                    m_Global.getFamilyList(0, cp)
                End If
            Else
                m_Global.getFamilyList(0, cp)
            End If
        End If

        m_Global.writeSurveyForm(0, htmlFlag, labelOption)
        prevFoundPersonList = cp
    End Sub

    Private Sub writeDirectoryBook()
        Dim cp As List(Of String) = readPersonIdFromDgChosen(0)
        If ugResult.Rows.Count = 0 Then
            m_Global.getFamilyList(0, cp)
        Else
            Dim str1 As String = ""
            Dim str2 As String = ""
            If cp.Count > 0 Then
                str1 = String.Join(",", cp)
            End If

            If prevFoundPersonList.Count > 0 Then
                str2 = String.Join(",", prevFoundPersonList)
            End If

            If ugResult.DisplayLayout.Bands(0).Columns("제외").Header.Caption = "제외" Then
                If str1 <> str2 Then
                    m_Global.getFamilyList(0, cp)
                End If
            Else
                m_Global.getFamilyList(0, cp)
            End If
        End If

        m_Global.writeDirectoryBook(0)
        prevFoundPersonList = cp
    End Sub

    Public Function readPersonIdFromDgChosen(ByVal colIndex As Integer) As List(Of String)
        readPersonIdFromDgChosen = New List(Of String)
        If dgChosen.RowCount > 0 Then
            For i = 0 To dgChosen.RowCount - 1
                readPersonIdFromDgChosen.Add(CType(dgChosen.Rows(i).Cells(colIndex).Value, String))
            Next
        End If
    End Function

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        frmAbout.ShowDialog()
    End Sub

    Private Sub tsbtExit_Click(sender As Object, e As EventArgs) Handles tsbtExit.Click
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub tsbtExcel_Click(sender As Object, e As EventArgs) Handles tsbtExcel.Click
        Me.Cursor = Cursors.WaitCursor
        Try
            Dim theFile As String = ROOT_FOLDER + "\" + m_Constant.TEMP_FILE 'Application.StartupPath + "\\" + BOOK_LIST_EXCEL
            Me.ugexcel.Export(Me.ugResult, theFile)
            Diagnostics.Process.Start(theFile)

        Catch ex As Exception
            MessageBox.Show(ex.Message + " 만일 " + ROOT_FOLDER + "에 " + m_Constant.TEMP_FILE + "이 열려 있으면 닫은 후에 다시 시도해주십시오.")
            Logger.LogError(ex.ToString)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub ugFind_DoubleClickRow(sender As Object, e As DoubleClickRowEventArgs) Handles ugFind.DoubleClickRow
        selectFoundPerson()
    End Sub

    Private Sub btFind_Click(sender As Object, e As EventArgs) Handles btFind.Click
        Dim sql As String = ""
        Me.Cursor = Cursors.WaitCursor
        Try
            If cbFindWhat.SelectedIndex = 1 Then
                If String.IsNullOrEmpty(txtFind.Text) = True Then
                    MsgBox("헌금봉투번호를 입력하고 다시 시도해주시기 바랍니다.", Nothing, "주의")
                    txtFind.Focus()
                    Me.Cursor = Cursors.Default
                    Exit Sub
                Else
                    If IsNumeric(txtFind.Text) = False Then
                        MsgBox("숫자를 입력하고 다시 시도해주시기 바랍니다.", Nothing, "주의")
                        txtFind.Clear()
                        txtFind.Focus()
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    End If
                End If
            ElseIf cbFindWhat.SelectedIndex = 5 Then
                If String.IsNullOrEmpty(txtFind.Text) = False Then
                    Dim fileHelper As FileHelper = New FileHelper
                    sql = fileHelper.readFromCSV(0, txtFind.Text)
                    If String.IsNullOrEmpty(sql) = True Then
                        MessageBox.Show("화일 읽기에 문제가 발생했습니다. 시스템 관리자에게 문의해주시기 바랍니다.", "경고")
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    End If
                Else
                    If openExcel() = False Then
                        Me.Cursor = Cursors.Default
                        Exit Sub
                    Else
                        If String.IsNullOrEmpty(sql) = True Then
                            Me.Cursor = Cursors.Default
                            Exit Sub
                        End If
                    End If
                End If
            ElseIf cbFindWhat.SelectedIndex = 6 Or
                   cbFindWhat.SelectedIndex = 8 Or
                   cbFindWhat.SelectedIndex = 10 Then '나이, 한글이름 중복, 주소갯수, 출생년도
                If String.IsNullOrEmpty(txtAgeFrom.Text) = True And String.IsNullOrEmpty(txtAgeTo.Text) = True Then
                    MessageBox.Show("연령대를 입력하거나 두개의 입력필드중 최소 하나의 숫자을 입력하여 주십시오.", "주의")
                    txtAgeFrom.Focus()
                    txtAgeFrom.SelectAll()
                    Me.Cursor = Cursors.Default
                    Exit Sub
                Else
                    If String.IsNullOrEmpty(txtAgeFrom.Text) = False And String.IsNullOrEmpty(txtAgeTo.Text) = False Then
                        Dim ageFrom As Integer = CType(txtAgeFrom.Text, Integer)
                        Dim ageTo As Integer = CType(txtAgeTo.Text, Integer)
                        If ageFrom > ageTo Then
                            MessageBox.Show("첫번째 숫자가 두번째 숫자보다 작은 수를 입력하여 주십시오.", "주의")
                            txtAgeFrom.Focus()
                            txtAgeFrom.SelectAll()
                            Me.Cursor = Cursors.Default
                            Exit Sub
                        Else
                            txtFind.Text = String.Empty
                        End If
                    End If
                End If
            ElseIf cbFindWhat.SelectedIndex = 11 Then
                If String.IsNullOrEmpty(txtAgeFrom.Text) = True And String.IsNullOrEmpty(txtAgeTo.Text) = True Then
                    MessageBox.Show("연령대를 입력하거나 두개의 입력필드중 최소 하나의 숫자을 입력하여 주십시오.", "주의")
                    txtAgeFrom.Focus()
                    txtAgeFrom.SelectAll()
                    Me.Cursor = Cursors.Default
                    Exit Sub
                Else
                    If String.IsNullOrEmpty(txtAgeFrom.Text) = False And String.IsNullOrEmpty(txtAgeTo.Text) = False Then
                        Dim ageFrom As Integer = CType(txtAgeFrom.Text, Integer)
                        Dim ageTo As Integer = CType(txtAgeTo.Text, Integer)
                        If ageFrom < ageTo Then
                            MessageBox.Show("첫번째 숫자가 두번째 숫자보다 큰 수를 입력하여 주십시오.", "주의")
                            txtAgeFrom.Focus()
                            txtAgeFrom.SelectAll()
                            Me.Cursor = Cursors.Default
                            Exit Sub
                        Else
                            txtFind.Text = String.Empty
                        End If
                    End If
                End If
            End If

            Dim isError As Boolean = findPerson(sql)
            txtFind.Enabled = True
            txtFind.Focus()
            txtFind.SelectAll()

            pnFind.Visible = Not isError
            ugFind.Visible = Not isError

        Catch ex As IOException
            Logger.LogError(ex.ToString)
            pnFind.Visible = False
            If ex.ToString.Contains("Could not find file") Then
                MsgBox("화일경로가 틀립니다. 다시 시도해주십시오.", Nothing, "주의")
                txtFind.Focus()
                txtFind.SelectAll()
            End If
        Catch ex As Exception
            Logger.LogError(ex.ToString)
            pnFind.Visible = False
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Function findPerson(ByVal namesSQL As String) As Boolean
        Me.Cursor = Cursors.WaitCursor
        Dim titleYear As String = ""
        findPerson = False

        If cbFindWhat.SelectedIndex = 3 Then
            Dim utilDAL As TKPC.DAL.UtilDAL = TKPC.DAL.UtilDAL.getInstance()
            Try
                Dim years As ArrayList = utilDAL.getTitleYears(cbType.SelectedIndex, txtFind.Text)
                titleYear = CType(years.Item(0), String)

                If years IsNot Nothing And years.Count > 0 Then
                    tscbYear.Items.Clear()
                    tscbYear.Items.Add("모든년도")
                    For i As Integer = 0 To years.Count - 1
                        tscbYear.Items.Add(years.Item(i))
                    Next
                End If

            Catch ex As Exception
                Logger.LogError(ex.ToString)
            Finally
                utilDAL = Nothing
            End Try
        End If

        Try
            findPerson = findPeopleList(namesSQL)

            If findPerson = False Then
                If cbFindWhat.SelectedIndex = 3 Then
                    If ugFind.Rows.Count > 0 Then
                        tslblYear.Visible = True
                        tscbYear.Visible = True
                        tscbYear.SelectedIndex = 0
                    Else
                        tslblYear.Visible = False
                        tscbYear.Visible = False
                    End If
                Else
                    tslblYear.Visible = False
                    tscbYear.Visible = False
                End If
            Else
                pnFind.Visible = False
            End If

        Catch ex As Exception
            Logger.LogError(ex.ToString)
            pnFind.Visible = False
        Finally
            Me.Cursor = Cursors.Default
        End Try
        Return findPerson
    End Function

    Private Sub cbFindWhat_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbFindWhat.SelectedIndexChanged
        txtFind.SelectAll()
        txtFind.Focus()

        '0 - 이름
        '1 - 헌금#
        '2 - 전화#
        '3 - 직분
        '4 - 주소
        '5 -엑셀-이름
        '6 - 나이
        '7 - 이멜
        '8 - 한글 이름 중복
        '9 - 독신
        '10 - 주소갯수
        '11 - 출생년도
        If cbFindWhat.SelectedIndex = 0 Then '이름
            cbType.Enabled = True
            cbType.SelectedIndex = 0
            txtFind.Visible = True
            PnAges.Visible = False
            txtFind.SelectAll()
            txtFind.Focus()

        ElseIf cbFindWhat.SelectedIndex = 1 Then '헌금번호
            If m_Global.accessLevel = m_Constant.ACCESS_LEVEL_PASTOR Then
                MsgBox("사용할 권한이 없습니다.", Nothing, "주의")
                cbFindWhat.SelectedIndex = 0
                Exit Sub
            End If
            cbType.Enabled = False
            cbType.SelectedIndex = 3
            txtFind.Visible = True
            PnAges.Visible = False
            txtFind.SelectAll()
            txtFind.Focus()

        ElseIf cbFindWhat.SelectedIndex = 2 Then '전화번호
            cbType.Enabled = True
            cbType.SelectedIndex = 3
            txtFind.Visible = True
            PnAges.Visible = False

        ElseIf cbFindWhat.SelectedIndex = 3 Then '직분명
            cbType.Enabled = True
            cbType.SelectedIndex = 3
            txtFind.Visible = True
            PnAges.Visible = False
            txtFind.SelectAll()
            txtFind.Focus()

        ElseIf cbFindWhat.SelectedIndex = 4 Then '주소
            cbType.Enabled = True
            cbType.SelectedIndex = 0
            txtFind.Visible = True
            PnAges.Visible = False
            txtFind.SelectAll()
            txtFind.Focus()

        ElseIf cbFindWhat.SelectedIndex = 5 Then '엑셀-이름
            txtFind.Visible = True
            PnAges.Visible = False
            If openExcel() Then
                Me.Cursor = Cursors.WaitCursor
                Dim fileHelper As FileHelper = New FileHelper
                Dim sql = fileHelper.readFromCSV(0, txtFind.Text)
                If String.IsNullOrEmpty(sql) = True Then
                    MessageBox.Show("화일 읽기에 문제가 발생했습니다.", "경고")
                    Me.Cursor = Cursors.Default
                    txtFind.Text = ""
                    txtFind.Enabled = True
                    Exit Sub
                Else
                    If sql = "-1" Then
                        Me.Cursor = Cursors.Default
                        txtFind.Text = ""
                        txtFind.Enabled = True
                        Exit Sub
                    End If
                End If
                findPerson(sql)
                txtFind.Enabled = True
                txtFind.Focus()
                txtFind.SelectAll()
                pnFind.Visible = True
                ugFind.Visible = True
                Me.Cursor = Cursors.Default
            End If

        ElseIf cbFindWhat.SelectedIndex = 6 Or
               cbFindWhat.SelectedIndex = 8 Or
               cbFindWhat.SelectedIndex = 10 Or
               cbFindWhat.SelectedIndex = 11 Then '나이, 이름 중복, 주소갯수, 출생년도
            cbType.Enabled = False
            txtFind.Visible = False
            PnAges.Visible = True
            txtAgeFrom.Focus()

        ElseIf cbFindWhat.SelectedIndex = 7 Then '이멜, 독신여부
            cbType.Enabled = True
            cbType.SelectedIndex = 0
            txtFind.Visible = True
            PnAges.Visible = False
            txtFind.SelectAll()
            txtFind.Focus()

        ElseIf cbFindWhat.SelectedIndex = 9 Then
            cbType.Enabled = False
            cbType.SelectedIndex = 3
            txtFind.Visible = True
            txtFind.Text = "y"
            PnAges.Visible = False
            txtFind.SelectAll()
            txtFind.Focus()
        End If
    End Sub

    Private Function openExcel() As Boolean
        openExcel = False
        cbType.Enabled = False
        cbType.SelectedIndex = 3
        Dim result As Integer = MessageBox.Show("주의사항" + vbCrLf + "1. UTF-8 Comma delimited CSV 또는 Unicode text 화일이어야 합니다." + vbCrLf + "2. 화일안의 첫번째 컬럼에 있는 교우 이름만 읽습니다." + vbCrLf + "3. 모든 이름은 붙여주십시오. " + vbCrLf + "4. 동명이인이 나올 수 있으니 반드시 확인해주시고 옮겨담기를 해주십시오." + vbCrLf + vbCrLf + "화일을 찾으시겠습니까?", "주의", MessageBoxButtons.YesNo)
        If result = DialogResult.Yes Then
            openExcel = True
            Try
                chkStatus3.Checked = True
                chkStatus4.Checked = True

                Dim OpenFileDialog As New OpenFileDialog
                OpenFileDialog.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
                OpenFileDialog.Filter = "Unicode txt files (*.txt)|*.txt|CSV Files (*.csv)|*.csv"
                If (OpenFileDialog.ShowDialog(Me) = System.Windows.Forms.DialogResult.OK) Then
                    txtFind.Text = OpenFileDialog.FileName
                    txtFind.Enabled = False
                End If
            Catch ex As Exception
                Logger.LogError(ex.ToString)
            Finally
            End Try
            Return openExcel
        Else
            Exit Function
        End If
    End Function

    Private Sub txtFind_KeyDown(sender As Object, e As KeyEventArgs) Handles txtFind.KeyDown
        If e.KeyCode = Keys.Enter Then
            btFind_Click(sender, e)
        ElseIf e.KeyCode = Keys.Down Then
            If ugFind.Visible = True Then
                ugFind.Focus()
                If ugFind.Rows.Count > 0 Then
                    ugFind.Rows(0).Selected = True
                End If
            End If
        End If
    End Sub

    Private Sub txtFind_TextChanged(sender As Object, e As EventArgs) Handles txtFind.TextChanged
        If String.IsNullOrEmpty(txtFind.Text) = True Then
            pnFind.Visible = False
        End If
    End Sub

    Private Sub ugFind_KeyDown(sender As Object, e As KeyEventArgs) Handles ugFind.KeyDown
        If e.KeyCode = Keys.Enter Then
            selectFoundPerson()
        ElseIf e.KeyCode = Keys.Escape Then
            txtFind.Text = ""
            txtFind.Focus()
            pnFind.Visible = False
        End If
    End Sub

    Public Sub selectFoundPerson()
        If ugFind.ActiveRow.Index <> -1 Then
            Dim personId As String = CType(ugFind.Rows(ugFind.ActiveRow.Index).Cells("person_id").Value, String)
            Dim koreanName As String = CType(ugFind.Rows(ugFind.ActiveRow.Index).Cells("korean_name").Value, String)
            Dim spouseIds As String = CType(ugFind.Rows(ugFind.ActiveRow.Index).Cells("spouses").Value, String)
            Dim status As String = CType(ugFind.Rows(ugFind.ActiveRow.Index).Cells("status").Value, String)
            Dim age As String = CType(ugFind.Rows(ugFind.ActiveRow.Index).Cells("age").Value, String)
            Dim gender As String = CType(ugFind.Rows(ugFind.ActiveRow.Index).Cells("gender").Value, String)

            If findPeopleHashTable.ContainsKey(personId) = False Then
                Dim personEnt As TKPC.Entity.PersonEnt = New TKPC.Entity.PersonEnt
                personEnt.personId = personId
                personEnt.koreanName = koreanName
                personEnt.spouseIds = spouseIds
                personEnt.status = status
                personEnt.age = age
                personEnt.gender = gender

                findPeopleHashTable.Add(personId, personEnt)
                With dgChosen
                    .Rows.Add(personId, koreanName, spouseIds, status, age, gender)
                    .Sort(dgChosen.Columns(1), System.ComponentModel.ListSortDirection.Ascending)
                End With
            End If

            With ugFind
                .DataSource = Nothing
                .Visible = False
            End With

            With txtFind
                .Clear()
                .Focus()
            End With
        End If
    End Sub

    Private Sub ugResult_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles ugResult.InitializeRow
        If familyMode = True Then
            'If e.Row.Cells("lvl").Text = "1-세대주" Then
            '    e.Row.Cells("제외").Activation = Infragistics.Win.UltraWinGrid.Activation.Disabled
            'End If
            Dim fileName As String = e.Row.Cells("PhotoName").Text
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

        If m_Global.accessLevel = m_Constant.ACCESS_LEVEL_OFFICE Then
            If cbCurrent.SelectedIndex = 1 Or cbCurrent.SelectedIndex = 5 Or cbCurrent.SelectedIndex = 6 Then
                If chkType.Checked = True Then
                    If e.Row.Cells("offering_type").Text = m_Constant.OFFERING_TOTAL Then
                        e.Row.CellAppearance.BackColor = Color.LightGreen
                    Else
                        e.Row.CellAppearance.BackColor = Color.White
                    End If
                End If

            ElseIf cbCurrent.SelectedIndex = 12 Or cbCurrent.SelectedIndex = 21 Or
                    cbCurrent.SelectedIndex = 23 Or cbCurrent.SelectedIndex = 24 Then
                If e.Row.Cells("lvl").Text = m_Constant.FAMILY_LEVEL_1 Then
                    e.Row.CellAppearance.BackColor = Color.LightGreen
                ElseIf e.Row.Cells("lvl").Text = m_Constant.FAMILY_LEVEL_0 Then
                    e.Row.CellAppearance.BackColor = Color.LightSalmon
                Else
                    e.Row.CellAppearance.BackColor = Color.White
                End If

            ElseIf cbCurrent.SelectedIndex = 9 Then '년도별 헌금봉투 찾기
                If String.IsNullOrEmpty(e.Row.Cells("person_id").Text) = True Then
                    e.Row.CellAppearance.BackColor = Color.LightSalmon
                Else
                    e.Row.CellAppearance.BackColor = Color.White
                End If

                If CType(e.Row.Cells("cnt").Text, Integer) > 2 Then
                    e.Row.CellAppearance.BackColor = Color.Red
                End If

            ElseIf cbCurrent.SelectedIndex = 10 Or cbCurrent.SelectedIndex = 11 Then
                If e.Row.Cells("남편_등록상태").Text <> e.Row.Cells("부인_등록상태").Text Then
                    e.Row.CellAppearance.BackColor = Color.LightSalmon
                End If

            ElseIf cbCurrent.SelectedIndex = 13 Or   '고인 리스트
               cbCurrent.SelectedIndex = 14 Or   '유아세례 리스트
               cbCurrent.SelectedIndex = 15 Or   '입교 리스트
               cbCurrent.SelectedIndex = 16 Or   '세례 리스트
               cbCurrent.SelectedIndex = 17 Or   '새교우 리스트
               cbCurrent.SelectedIndex = 18 Or   '등록 리스트
               cbCurrent.SelectedIndex = 19 Or   '이적 리스트 
               cbCurrent.SelectedIndex = 20 Then '생일 리스트
                If familyMode = True Then
                    If e.Row.Cells("lvl").Text = m_Constant.FAMILY_LEVEL_1 Then
                        e.Row.CellAppearance.BackColor = Color.LightGreen
                    ElseIf e.Row.Cells("lvl").Text = m_Constant.FAMILY_LEVEL_0 Then
                        e.Row.CellAppearance.BackColor = Color.LightSalmon
                    Else
                        e.Row.CellAppearance.BackColor = Color.White
                    End If
                Else

                End If
            End If

        Else
            If cbPastor.SelectedIndex = 0 Or cbPastor.SelectedIndex = 4 Or cbPastor.SelectedIndex = 5 Then
                If chkType.Checked = True Then
                    If e.Row.Cells("offering_type").Text = m_Constant.OFFERING_TOTAL Then
                        e.Row.CellAppearance.BackColor = Color.LightGreen
                    Else
                        e.Row.CellAppearance.BackColor = Color.White
                    End If
                End If

            ElseIf cbPastor.SelectedIndex = 9 Or cbPastor.SelectedIndex = 18 Or
                    cbPastor.SelectedIndex = 20 Then
                If e.Row.Cells("lvl").Text = m_Constant.FAMILY_LEVEL_1 Then
                    e.Row.CellAppearance.BackColor = Color.LightGreen
                ElseIf e.Row.Cells("lvl").Text = m_Constant.FAMILY_LEVEL_0 Then
                    e.Row.CellAppearance.BackColor = Color.LightSalmon
                Else
                    e.Row.CellAppearance.BackColor = Color.White
                End If

            ElseIf cbPastor.SelectedIndex = 7 Or cbPastor.SelectedIndex = 8 Then
                If e.Row.Cells("남편_등록상태").Text <> e.Row.Cells("부인_등록상태").Text Then
                    e.Row.CellAppearance.BackColor = Color.LightSalmon
                End If

            ElseIf cbPastor.SelectedIndex = 10 Or   '고인 리스트
               cbPastor.SelectedIndex = 11 Or   '유아세례 리스트
               cbPastor.SelectedIndex = 12 Or   '입교 리스트
               cbPastor.SelectedIndex = 13 Or   '세례 리스트
               cbPastor.SelectedIndex = 14 Or   '새교우 리스트
               cbPastor.SelectedIndex = 15 Or   '등록 리스트
               cbPastor.SelectedIndex = 16 Or   '이적 리스트 
               cbPastor.SelectedIndex = 17 Then '생일 리스트
                If familyMode = True Then
                    If e.Row.Cells("lvl").Text = m_Constant.FAMILY_LEVEL_1 Then
                        e.Row.CellAppearance.BackColor = Color.LightGreen
                    ElseIf e.Row.Cells("lvl").Text = m_Constant.FAMILY_LEVEL_0 Then
                        e.Row.CellAppearance.BackColor = Color.LightSalmon
                    Else
                        e.Row.CellAppearance.BackColor = Color.White
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub ugResult_InitializeLayout(sender As Object, e As InitializeLayoutEventArgs) Handles ugResult.InitializeLayout
        ugResult.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.True

        If familyMode = True Then
            Dim checkColumn As UltraGridColumn = e.Layout.Bands(0).Columns.Add("제외", "제외")
            checkColumn.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.CheckBox
            checkColumn.DataType = Type.GetType("System.Boolean")
            checkColumn.DefaultCellValue = False
            checkColumn.CellActivation = Activation.AllowEdit
            checkColumn.Header.VisiblePosition = 0
            checkColumn.CellClickAction = CellClickAction.Edit
            checkColumn.Width = 60

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

            e.Layout.Override.HeaderClickAction = HeaderClickAction.Select
        End If

        If m_Global.accessLevel = m_Constant.ACCESS_LEVEL_OFFICE Then
            If cbCurrent.SelectedIndex = 1 Then
                If chkAll.Checked = False Then
                    e.Layout.Override.HeaderClickAction = HeaderClickAction.Select
                Else
                    e.Layout.Override.HeaderClickAction = HeaderClickAction.SortMulti
                End If

            ElseIf cbCurrent.SelectedIndex = 9 Then '년도별 헌금봉투 찾기
                e.Layout.Bands(0).Columns("person_id").Hidden = True
                e.Layout.Bands(0).Columns("cnt").Hidden = True

            ElseIf cbCurrent.SelectedIndex = 12 Or cbCurrent.SelectedIndex = 21 Then
                e.Layout.Override.HeaderClickAction = HeaderClickAction.Select

            ElseIf cbCurrent.SelectedIndex = 13 Or   '고인 리스트
                   cbCurrent.SelectedIndex = 14 Or   '유아세례 리스트
                   cbCurrent.SelectedIndex = 15 Or   '입교 리스트
                   cbCurrent.SelectedIndex = 16 Or   '세례 리스트
                   cbCurrent.SelectedIndex = 17 Or   '새교우 리스트
                   cbCurrent.SelectedIndex = 18 Or   '등록 리스트
                   cbCurrent.SelectedIndex = 19 Or   '이적 리스트
                   cbCurrent.SelectedIndex = 20 Then '생일 리스트

                If familyMode = False Then
                    e.Layout.Bands(0).Columns("photo_file_name").Hidden = True
                    e.Layout.Bands(0).Columns("spouses").Hidden = True
                    e.Layout.Bands(0).Columns("person_id").Hidden = True
                    'e.Layout.Bands(0).Columns("stay").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Time
                Else
                    e.Layout.Bands(0).Columns("PhotoName").Hidden = True
                    e.Layout.Bands(0).Columns("id").Hidden = True
                End If


                If chkAll.Checked = True Then
                    e.Layout.Override.HeaderClickAction = HeaderClickAction.Select
                Else
                    e.Layout.Override.HeaderClickAction = HeaderClickAction.SortMulti
                End If

            ElseIf cbCurrent.SelectedIndex = 22 Then
                e.Layout.Bands(0).Columns("person_id").Hidden = True
            End If
        Else '목회자 모드
            If cbPastor.SelectedIndex = 0 Then
                If chkAll.Checked = False Then
                    e.Layout.Override.HeaderClickAction = HeaderClickAction.Select
                Else
                    e.Layout.Override.HeaderClickAction = HeaderClickAction.SortMulti
                End If

            ElseIf cbPastor.SelectedIndex = 9 Or cbPastor.SelectedIndex = 18 Then
                e.Layout.Override.HeaderClickAction = HeaderClickAction.Select

            ElseIf cbPastor.SelectedIndex = 10 Or   '고인 리스트
                   cbPastor.SelectedIndex = 11 Or   '유아세례 리스트
                   cbPastor.SelectedIndex = 12 Or   '입교 리스트
                   cbPastor.SelectedIndex = 13 Or   '세례 리스트
                   cbPastor.SelectedIndex = 14 Or   '새교우 리스트
                   cbPastor.SelectedIndex = 15 Or   '등록 리스트
                   cbPastor.SelectedIndex = 16 Or   '이적 리스트
                   cbPastor.SelectedIndex = 17 Then '생일 리스트

                If familyMode = False Then
                    e.Layout.Bands(0).Columns("photo_file_name").Hidden = True
                    e.Layout.Bands(0).Columns("spouses").Hidden = True
                    e.Layout.Bands(0).Columns("person_id").Hidden = True
                    'e.Layout.Bands(0).Columns("stay").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Time
                Else
                    e.Layout.Bands(0).Columns("PhotoName").Hidden = True
                    e.Layout.Bands(0).Columns("id").Hidden = True
                End If

                'e.Layout.Bands(0).Columns("stay").Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Time

                If chkAll.Checked = True Then
                    e.Layout.Override.HeaderClickAction = HeaderClickAction.Select
                Else
                    e.Layout.Override.HeaderClickAction = HeaderClickAction.SortMulti
                End If
            ElseIf cbPastor.SelectedIndex = 19 Then
                e.Layout.Bands(0).Columns("person_id").Hidden = True
            End If
        End If
    End Sub

    Private Sub ugResult_MouseClick(sender As Object, e As MouseEventArgs) Handles ugResult.MouseClick
        If pnFind.Visible = True Then
            pnFind.Visible = False
        End If
    End Sub

    Private Sub pnCurrent1_MouseClick(sender As Object, e As MouseEventArgs) Handles pnCurrent1.MouseClick
        If pnFind.Visible = True Then
            pnFind.Visible = False
        End If
    End Sub

    Private Sub pTop_MouseClick(sender As Object, e As MouseEventArgs) Handles pTop.MouseClick
        If pnFind.Visible = True Then
            pnFind.Visible = False
        End If
    End Sub

    'Private Sub tsbtCard_Click(sender As Object, e As EventArgs) Handles tsbtCard.Click
    '    writeMembershipCard()
    'End Sub

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

    Private Sub txtFind_GotFocus(sender As Object, e As EventArgs) Handles txtFind.GotFocus
        If whichIME = 0 Then
            whichIME = SwitchIME(hKrLayoutId, KbdKr)
            'My.Computer.Keyboard.SendKeys("{%}")
            'SendKeys.Send("%")
        End If
    End Sub

    Private Sub ugFind_InitializeLayout(sender As Object, e As InitializeLayoutEventArgs) Handles ugFind.InitializeLayout
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

        Dim btnColumn2 As UltraGridColumn = e.Layout.Bands(0).Columns.Add("map", "map")
        btnColumn2.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button
        btnColumn2.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always
        'btnColumn.CellButtonAppearance.Image = "C:\Users\hylee\Documents\TKPC\system\VBExpress\TKPCAdminReportMaker\TKPCAdminReportMaker\TKPCAdminReportMaker\Resources\view_picture.png"
        btnColumn2.CellButtonAppearance.TextHAlign = Infragistics.Win.HAlign.Center
        btnColumn2.Width = 50
        btnColumn2.DefaultCellValue = False
        btnColumn2.CellActivation = Activation.AllowEdit
        btnColumn2.Header.VisiblePosition = 2
        btnColumn2.CellClickAction = CellClickAction.Edit

        Dim btnColumn3 As UltraGridColumn = e.Layout.Bands(0).Columns.Add("dir", "dir")
        btnColumn3.Style = Infragistics.Win.UltraWinGrid.ColumnStyle.Button
        btnColumn3.ButtonDisplayStyle = Infragistics.Win.UltraWinGrid.ButtonDisplayStyle.Always
        'btnColumn.CellButtonAppearance.Image = "C:\Users\hylee\Documents\TKPC\system\VBExpress\TKPCAdminReportMaker\TKPCAdminReportMaker\TKPCAdminReportMaker\Resources\view_picture.png"
        btnColumn3.CellButtonAppearance.TextHAlign = Infragistics.Win.HAlign.Center
        btnColumn3.Width = 50
        btnColumn3.DefaultCellValue = False
        btnColumn3.CellActivation = Activation.AllowEdit
        btnColumn3.Header.VisiblePosition = 3
        btnColumn3.CellClickAction = CellClickAction.Edit

        If cbFindWhat.SelectedIndex = 1 Then '헌금봉투번호
            With ugFind
                .DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True
                .DisplayLayout.Bands(0).Columns(0).Width = 60 'person_id
                .DisplayLayout.Bands(0).Columns(0).Hidden = True 'person_id
                .DisplayLayout.Bands(0).Columns(1).Width = 50 'year
                .DisplayLayout.Bands(0).Columns(2).Width = 70 'Korean Name
                .DisplayLayout.Bands(0).Columns(3).Width = 50 'Korean Name count
                .DisplayLayout.Bands(0).Columns(4).Width = 70 'title
                .DisplayLayout.Bands(0).Columns(5).Width = 70 'title date
                .DisplayLayout.Bands(0).Columns(6).Width = 70 'title years
                .DisplayLayout.Bands(0).Columns(7).Width = 90 'offering years
                .DisplayLayout.Bands(0).Columns(8).Width = 70 'Last Name
                .DisplayLayout.Bands(0).Columns(9).Width = 100 'First Name
                .DisplayLayout.Bands(0).Columns(10).Width = 30 'Gender
                .DisplayLayout.Bands(0).Columns(11).Width = 70 'DOB
                .DisplayLayout.Bands(0).Columns(12).Width = 40 'DOB year
                .DisplayLayout.Bands(0).Columns(13).Width = 35 'Age
                .DisplayLayout.Bands(0).Columns(14).Width = 45 'single
                .DisplayLayout.Bands(0).Columns(15).Width = 70 'Age range
                .DisplayLayout.Bands(0).Columns(16).Width = 120 'Email
                .DisplayLayout.Bands(0).Columns(17).Width = 80 'Cell Phone
                .DisplayLayout.Bands(0).Columns(18).Width = 80 'home Phone
                .DisplayLayout.Bands(0).Columns(19).Width = 80 'work Phone
                .DisplayLayout.Bands(0).Columns(20).Width = 295 'Home address
                .DisplayLayout.Bands(0).Columns(21).Width = 60 'address cnt
                .DisplayLayout.Bands(0).Columns(22).Width = 50 'status
                .DisplayLayout.Bands(0).Columns(23).Width = 90 'status_date
                .DisplayLayout.Bands(0).Columns(24).Width = 90 'deceased_on
                .DisplayLayout.Bands(0).Columns(25).Width = 80 'spouses
                .DisplayLayout.Bands(0).Columns(25).Hidden = True
                .DisplayLayout.Bands(0).Columns(26).Hidden = True 'photo_file_name
                .DisplayLayout.Override.RowAppearance.TextVAlign = Infragistics.Win.VAlign.Middle
            End With
        Else '이름, 전화#, 직분명, 도시명
            With ugFind
                .DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.True
                .DisplayLayout.Bands(0).Columns(0).Width = 60 'person_id
                .DisplayLayout.Bands(0).Columns(0).Hidden = True 'person_id
                .DisplayLayout.Bands(0).Columns(1).Width = 70 'Korean Name
                .DisplayLayout.Bands(0).Columns(2).Width = 50 'Korean Name count
                .DisplayLayout.Bands(0).Columns(3).Width = 70 'title
                .DisplayLayout.Bands(0).Columns(4).Width = 70 'title date
                .DisplayLayout.Bands(0).Columns(5).Width = 70 'title years
                .DisplayLayout.Bands(0).Columns(6).Width = 90 'offering years
                .DisplayLayout.Bands(0).Columns(7).Width = 70 'Last Name
                .DisplayLayout.Bands(0).Columns(8).Width = 100 'First Name
                .DisplayLayout.Bands(0).Columns(9).Width = 30 'Gender
                .DisplayLayout.Bands(0).Columns(10).Width = 70 'DOB
                .DisplayLayout.Bands(0).Columns(11).Width = 40 'DOB year
                .DisplayLayout.Bands(0).Columns(12).Width = 35 'Age
                .DisplayLayout.Bands(0).Columns(13).Width = 45 'single
                .DisplayLayout.Bands(0).Columns(14).Width = 70 'Age range
                .DisplayLayout.Bands(0).Columns(15).Width = 120 'Email
                .DisplayLayout.Bands(0).Columns(16).Width = 80 'Cell Phone
                .DisplayLayout.Bands(0).Columns(17).Width = 80 'home Phone
                .DisplayLayout.Bands(0).Columns(18).Width = 80 'work Phone
                .DisplayLayout.Bands(0).Columns(19).Width = 295 'Home address
                .DisplayLayout.Bands(0).Columns(20).Width = 60 'address cnt
                .DisplayLayout.Bands(0).Columns(21).Width = 50 'status
                .DisplayLayout.Bands(0).Columns(22).Width = 90 'status_date
                .DisplayLayout.Bands(0).Columns(23).Width = 90 'deceased_on
                .DisplayLayout.Bands(0).Columns(24).Width = 80 'spouses
                .DisplayLayout.Bands(0).Columns(24).Hidden = True
                .DisplayLayout.Bands(0).Columns(25).Hidden = True 'photo_file_name
                .DisplayLayout.Override.RowAppearance.TextVAlign = Infragistics.Win.VAlign.Middle
            End With
        End If

        findPeopleRowSelected = 0
    End Sub

    Private Sub ugFind_InitializeRow(sender As Object, e As InitializeRowEventArgs) Handles ugFind.InitializeRow
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

        Dim address As String = e.Row.Cells("address").Text
        If String.IsNullOrEmpty(address) = False Then
            e.Row.Cells("map").Value = "map"
        Else
            e.Row.Cells("map").Value = "update"
            e.Row.Cells("map").Activation = Infragistics.Win.UltraWinGrid.Activation.Disabled
        End If

        If String.IsNullOrEmpty(address) = False Then
            e.Row.Cells("dir").Value = "go"
        Else
            e.Row.Cells("dir").Value = "update"
            e.Row.Cells("dir").Activation = Infragistics.Win.UltraWinGrid.Activation.Disabled
        End If

        Dim deceased As String = e.Row.Cells("deceased_on").Text
        If String.IsNullOrEmpty(deceased) = False Then
            e.Row.Cells("status").Value = "고인"
        End If

        If findPeopleHashTable.ContainsKey(e.Row.Cells("person_id").Text) = True Then
            e.Row.Cells("choose").Value = "True"
            e.Row.Appearance.BackColor = Color.LightPink
        End If

        If String.IsNullOrEmpty(e.Row.Cells("age").Text) = True Then
            e.Row.Cells("age").Value = -1
        End If

        'If e.Row.Cells("choose").Text = "True" Then
        '    findPeopleRowSelected += 1
        'Else
        '    findPeopleRowSelected -= 1
        '    If findPeopleRowSelected <0 Then
        '        findPeopleRowSelected = 0
        '    End If
        'End If
    End Sub

    Private Sub ugFind_CellChange(sender As Object, e As CellEventArgs) Handles ugFind.CellChange
        ugFind.DisplayLayout.CaptionVisible = Infragistics.Win.DefaultableBoolean.True

        If e.Cell.Text = "True" Then
            findPeopleRowSelected += 1
            e.Cell.Row.Appearance.BackColor = Color.LightPink
        Else
            findPeopleRowSelected -= 1
            If findPeopleRowSelected < 0 Then
                findPeopleRowSelected = 0
            End If

            e.Cell.Row.Appearance.BackColor = Color.White
        End If

        If findPeopleRowSelected > 0 Then
            tsbtFindMove.Enabled = True
        Else
            tsbtFindMove.Enabled = False
        End If

        'MsgBox(findPeopleRowSelected)
    End Sub

    'Private Sub chkAll_CheckedChanged(sender As Object, e As EventArgs) Handles chkAll.CheckedChanged
    '    If cbCurrent.SelectedIndex = 13 Or   '고인 리스트
    '           cbCurrent.SelectedIndex = 14 Or   '유아세례 리스트
    '           cbCurrent.SelectedIndex = 15 Or   '입교 리스트
    '           cbCurrent.SelectedIndex = 16 Or   '세례 리스트
    '           cbCurrent.SelectedIndex = 17 Or   '새교우 리스트
    '           cbCurrent.SelectedIndex = 18 Or   '등록 리스트
    '           cbCurrent.SelectedIndex = 19 Then '생일 리스트
    '        If chkAll.Checked = True Then
    '            If dgChosen.RowCount > 0 Then
    '                MsgBox("전체 개인 옵션을 선택했습니다. 선택한 사람들을 모두 삭제한 후 다시 시도하십시오.", Nothing, "주의")
    '                chkAll.Checked = False
    '            End If
    '        End If
    '    End If
    'End Sub

    Private Sub ugFind_ClickCellButton(sender As Object, e As CellEventArgs) Handles ugFind.ClickCellButton
        If e.Cell.Text = "map" Then
            Dim address As String = e.Cell.Row.Cells("address").Text
            Dim url As String = "https://www.google.ca/maps/place/" + address.Replace(" ", "+")
            System.Diagnostics.Process.Start(url)
        ElseIf e.Cell.Text = "go" Then
            Dim address As String = e.Cell.Row.Cells("address").Text
            Dim url As String = "https://www.google.ca/maps/dir/" + m_Constant.ADDRESS_TKPC.Replace(" ", "+") + "/" + address.Replace(" ", "+")
            System.Diagnostics.Process.Start(url)
        Else
            With frmDetail
                .personId = e.Cell.Row.Cells("person_id").Text

                If e.Cell.Row.Cells("picture").Text = "view" Then
                    .photo_file_name = m_Constant.PERSON_IMAGE_FOLDER + e.Cell.Row.Cells("photo_file_name").Text
                Else
                    .photo_file_name = m_Constant.PERSON_IMAGE_FOLDER + "no_image.jpg"
                End If
                .ShowDialog()
            End With
        End If
    End Sub

    Private Sub tsbtFindAll_Click(sender As Object, e As EventArgs) Handles tsbtFindAll.Click
        With ugFind
            For i As Integer = 0 To .Rows.Count - 1
                .Rows(i).Cells("choose").Value = True
                .Rows(i).Appearance.BackColor = Color.LightPink
            Next
        End With
        tsbtFindMove.Enabled = True
        findPeopleRowSelected = ugFind.Rows.Count
    End Sub

    Private Sub tsbtFindUnAll_Click(sender As Object, e As EventArgs) Handles tsbtFindUnAll.Click
        With ugFind
            For i As Integer = 0 To .Rows.Count - 1
                .Rows(i).Cells("choose").Value = False
                .Rows(i).Appearance.BackColor = Color.White
            Next
        End With
        tsbtFindMove.Enabled = False
        findPeopleRowSelected = 0
    End Sub

    Private Sub tsbtFindMove_Click(sender As Object, e As EventArgs) Handles tsbtFindMove.Click
        If ugFind.Rows.Count > 0 Then

            For i As Integer = 0 To ugFind.Rows.Count - 1
                Dim personId As String = CType(ugFind.Rows(i).Cells("person_id").Value, String)
                Dim koreanName As String = CType(IIf(IsDBNull(ugFind.Rows(i).Cells("korean_name").Value), "", ugFind.Rows(i).Cells("korean_name").Value), String)
                Dim spouseIds As String = CType(ugFind.Rows(i).Cells("spouses").Value, String)
                Dim status As String = CType(ugFind.Rows(i).Cells("status").Value, String)
                Dim age As String = CType(ugFind.Rows(i).Cells("age").Value, String)
                Dim gender As String = CType(ugFind.Rows(i).Cells("gender").Value, String)

                If ugFind.Rows(i).Cells("choose").Text = "True" Then
                    If findPeopleHashTable.ContainsKey(personId) = False Then
                        Dim personEnt As TKPC.Entity.PersonEnt = New TKPC.Entity.PersonEnt
                        personEnt.personId = personId
                        personEnt.koreanName = koreanName
                        personEnt.spouseIds = spouseIds
                        personEnt.status = status
                        personEnt.age = age
                        personEnt.gender = gender

                        findPeopleHashTable.Add(personId, personEnt)
                    End If
                Else
                    '헌금으로 찾기를 할 경우 똑같은 사람이 나온다. 
                    '이때 첵크박스를 체크하고 하지 않은 로우가 있을 때 체크한 로우가 삭제된다.
                    If findPeopleHashTable.ContainsKey(personId) = True And cbFindWhat.SelectedIndex <> 1 Then
                        findPeopleHashTable.Remove(personId)
                    End If
                End If
            Next

            If findPeopleHashTable.Count > 0 Then
                dgChosen.Rows.Clear()
                Me.Cursor = Cursors.WaitCursor
                For Each dic As DictionaryEntry In findPeopleHashTable
                    Dim personEnt As TKPC.Entity.PersonEnt = CType(dic.Value, TKPC.Entity.PersonEnt)

                    With dgChosen
                        .Rows.Add(personEnt.personId, personEnt.koreanName, personEnt.spouseIds, personEnt.status, personEnt.age, personEnt.gender)
                        .Sort(dgChosen.Columns(1), System.ComponentModel.ListSortDirection.Ascending)
                    End With
                Next
                Me.Cursor = Cursors.Default
                txtFind.SelectAll()
                txtFind.Focus()
            End If
        End If
    End Sub

    Private Sub tsFindExcelMenu1_Click(sender As Object, e As EventArgs) Handles tsFindExcelMenu1.Click
        Me.Cursor = Cursors.WaitCursor
        Try
            Dim theFile As String = FILE_FOLDER + "\" + m_Constant.FOUND_PEOPLE_FILE 'Application.StartupPath + "\\" + BOOK_LIST_EXCEL
            Me.ugexcel.Export(Me.ugFind, theFile)
            Diagnostics.Process.Start(theFile)

        Catch ex As Exception
            MessageBox.Show(ex.Message + " 만일 " + FILE_FOLDER + "found_people_list.xls 가 열려 있다면, 닫고 다시 시도해주시기 바랍니다.")
            Logger.LogError(ex.ToString)
        Finally
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub tsFindExcelMenu2_Click(sender As Object, e As EventArgs) Handles tsFindExcelMenu2.Click
        Dim row As UltraGridRow
        For Each row In ugFind.Rows
            If row.Cells("choose").Text = "True" Then
                row.Selected = True
                row.Appearance.BackColor = Color.White
                ugFind.GetRowFromPrintRow(row)
                row.Hidden = False
            Else
                row.Selected = False
                row.Hidden = True
            End If
        Next

        Me.Cursor = Cursors.WaitCursor
        Try
            Dim theFile As String = FILE_FOLDER + "\" + m_Constant.FOUND_PEOPLE_FILE 'Application.StartupPath + "\\" + BOOK_LIST_EXCEL
            Me.ugexcel.Export(Me.ugFind, theFile)
            Diagnostics.Process.Start(theFile)

        Catch ex As Exception
            MessageBox.Show(ex.Message + " 만일 " + FILE_FOLDER + "found_people_list.xls 가 열려 있다면, 닫고 다시 시도해주시기 바랍니다.")
            Logger.LogError(ex.ToString)
        Finally
            Me.Cursor = Cursors.Default
        End Try

        For Each row In ugFind.Rows
            row.Hidden = False
            If row.Cells("choose").Text = "True" Then
                row.Appearance.BackColor = Color.LightPink
            End If
        Next
    End Sub

    Private Sub tsFindExcelMenu3_Click(sender As Object, e As EventArgs) Handles tsFindExcelMenu3.Click
        Dim row As UltraGridRow
        For Each row In ugFind.Rows
            If row.Cells("choose").Text = "False" Then
                row.Selected = True
                ugFind.GetRowFromPrintRow(row)
                row.Hidden = False
            Else
                row.Selected = False
                row.Hidden = True
            End If
        Next

        Me.Cursor = Cursors.WaitCursor
        Try
            Dim theFile As String = FILE_FOLDER + "\" + m_Constant.FOUND_PEOPLE_FILE 'Application.StartupPath + "\\" + BOOK_LIST_EXCEL
            Me.ugexcel.Export(Me.ugFind, theFile)
            Diagnostics.Process.Start(theFile)

        Catch ex As Exception
            MessageBox.Show(ex.Message + " 만일 " + FILE_FOLDER + "found_people_list.xls 가 열려 있다면, 닫고 다시 시도해주시기 바랍니다.")
            Logger.LogError(ex.ToString)
        Finally
            Me.Cursor = Cursors.Default
        End Try

        For Each row In ugFind.Rows
            row.Hidden = False
        Next
    End Sub

    Private Sub tsbtFindExit_Click(sender As Object, e As EventArgs) Handles tsbtFindExit.Click
        ugFind.DataSource = Nothing
        pnFind.Visible = False
    End Sub

    Private Sub mdiTKPC_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Dim fileHelper As FileHelper = New FileHelper
        fileHelper.deleteFiles(7, m_Constant.FILE_EXTENSION_LOG)
    End Sub

    Private Sub tsbtSwitch_Click(sender As Object, e As EventArgs)
        If MsgBox("가족리스트를 위해 세대주를 변경시킵니다. 진행하시겠습니까? ", MsgBoxStyle.YesNo, "선택") = vbYes Then
            Dim personDAL As TKPC.DAL.PersonDAL = TKPC.DAL.PersonDAL.getInstance()
            Dim restoreSpousePosition As String = ""
            Dim countOfspousePositionChanged As Integer = 0

            restoreSpousePosition = personDAL.restoreUpdatedSpousePositions()
            Logger.LogInfo(restoreSpousePosition)
            countOfspousePositionChanged = personDAL.updateSpousePosition()
            Logger.LogInfo("The number of changed spouses' positions: " + CType(countOfspousePositionChanged, String))

            MsgBox("총 " + CType(countOfspousePositionChanged, String) + "쌍의 세대주가 변경되었습니다.", Nothing, "결과")

            personDAL = Nothing
        End If
    End Sub

    Private Sub 직분ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles tsmTitle.Click
        Dim ht As Hashtable = New Hashtable
        With dgChosen
            For i As Integer = 0 To .Rows.Count - 1
                Dim personEnt As TKPC.Entity.PersonEnt = New TKPC.Entity.PersonEnt
                personEnt.personId = CType(dgChosen.Rows(i).Cells("person_id").Value, String)
                personEnt.koreanName = CType(dgChosen.Rows(i).Cells("korean_name").Value, String)
                personEnt.age = CType(dgChosen.Rows(i).Cells("age").Value, String)
                personEnt.gender = CType(dgChosen.Rows(i).Cells("gender").Value, String)
                personEnt.status = CType(dgChosen.Rows(i).Cells("status").Value, String)

                ht.Add(personEnt.personId, personEnt)
                personEnt = Nothing
            Next
        End With

        With frmUpdateAllOption
            .dataHT = ht
            .milestoneType = "TitleMilestone"
            .ShowDialog()
        End With
    End Sub

    Private Sub 신급ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles tsmBaptism.Click
        Dim ht As Hashtable = New Hashtable
        With dgChosen
            For i As Integer = 0 To .Rows.Count - 1
                Dim personEnt As TKPC.Entity.PersonEnt = New TKPC.Entity.PersonEnt
                personEnt.personId = CType(dgChosen.Rows(i).Cells("person_id").Value, String)
                personEnt.koreanName = CType(dgChosen.Rows(i).Cells("korean_name").Value, String)
                personEnt.age = CType(dgChosen.Rows(i).Cells("age").Value, String)
                personEnt.gender = CType(dgChosen.Rows(i).Cells("gender").Value, String)
                personEnt.status = CType(dgChosen.Rows(i).Cells("status").Value, String)

                ht.Add(personEnt.personId, personEnt)
                personEnt = Nothing
            Next
        End With

        With frmUpdateAllOption
            .dataHT = ht
            .milestoneType = "BaptismMilestone"
            .ShowDialog()
        End With
    End Sub

    Private Sub tsbtRemoveOne_Click(sender As Object, e As EventArgs) Handles tsbtRemoveOne.Click
        If dgChosen.RowCount = 0 Then Exit Sub
        Dim result As Integer = MessageBox.Show("선택한 레코드를 삭제하시겠습니까?", "삭제", MessageBoxButtons.YesNo)
        If result = DialogResult.Yes Then
            Dim personId As String = CType(dgChosen.Rows(dgChosen.CurrentRow.Index).Cells(0).Value, String)
            If findPeopleHashTable.ContainsKey(personId) Then
                findPeopleHashTable.Remove(personId)
            End If
            dgChosen.Rows.RemoveAt(dgChosen.CurrentRow.Index)

            If dgChosen.Rows.Count = 0 Then
                enableUpdateAll(False, False, False, False, False)
            Else
                enableUpdateAll(True, True, True, True, True)
            End If
        End If
        txtFind.Focus()
    End Sub

    Private Sub tsbtRemoveAll_Click(sender As Object, e As EventArgs) Handles tsbtRemoveAll.Click
        If dgChosen.RowCount = 0 Then Exit Sub
        Dim result As Integer = MessageBox.Show("레코드를 모두 삭제하시겠습니까?", "모두 삭제", MessageBoxButtons.YesNo)
        If result = DialogResult.Yes Then
            dgChosen.Rows.Clear()
            findPeopleHashTable.Clear()
            enableUpdateAll(False, False, False, False, False)
            If pnFind.Visible = True Then
                tsbtFindUnAll_Click(sender, e)
            End If
        End If
    End Sub

    Private Sub enableUpdateAll(ByVal flag As Boolean, ByVal flag1 As Boolean, ByVal flag2 As Boolean, ByVal flag3 As Boolean, ByVal flag4 As Boolean)
        tssbtUpdateAll.Enabled = flag
        tsmBaptism.Enabled = flag1
        tsmTitle.Enabled = flag2
        tsbtRemoveOne.Enabled = flag3
        tsbtRemoveAll.Enabled = flag4
    End Sub

    Private Sub 모든데이터ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 모든데이터ToolStripMenuItem.Click
        Dim names As String = ""
        With ugFind
            For i As Integer = 0 To .Rows.Count - 1
                Dim koreanName As String = ugFind.Rows(i).Cells("korean_name").Text
                If String.IsNullOrEmpty(koreanName) = False Then
                    names += koreanName
                    If i < .Rows.Count - 1 Then
                        names += ","
                    End If
                End If
            Next
        End With

        With frmTextNames
            .names = names
            .ShowDialog()
        End With
    End Sub

    Private Sub 선택하지않은데이터만ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 선택하지않은데이터만ToolStripMenuItem.Click
        Dim names As String = ""
        With ugFind
            For i As Integer = 0 To .Rows.Count - 1
                If ugFind.Rows(i).Cells("choose").Text = "False" Then
                    Dim koreanName As String = ugFind.Rows(i).Cells("korean_name").Text
                    If String.IsNullOrEmpty(koreanName) = False Then
                        names += koreanName
                        If i < .Rows.Count - 1 Then
                            names += ","
                        End If
                    End If
                End If
            Next
        End With

        With frmTextNames
            .names = names
            .ShowDialog()
        End With
    End Sub

    Private Sub 선택한데이터만ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 선택한데이터만ToolStripMenuItem.Click
        Dim names As String = ""
        With ugFind
            For i As Integer = 0 To .Rows.Count - 1
                If ugFind.Rows(i).Cells("choose").Text = "True" Then
                    Dim koreanName As String = ugFind.Rows(i).Cells("korean_name").Text
                    If String.IsNullOrEmpty(koreanName) = False Then
                        names += koreanName
                        If i < .Rows.Count - 1 Then
                            names += ","
                        End If
                    End If
                End If
            Next
        End With

        With frmTextNames
            .names = names
            .ShowDialog()
        End With
    End Sub

    Private Sub 세대주변경ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 세대주변경ToolStripMenuItem.Click
        Me.Cursor = Cursors.WaitCursor
        frmSwitchFamilyHead.ShowDialog()
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub SetMaximumOfferingNumberToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SetMaximumOfferingNumberToolStripMenuItem.Click
        frmMaxOfferingNumber.ShowDialog()
    End Sub

    Private Sub tscbYear_DropDownClosed(sender As Object, e As EventArgs) Handles tscbYear.DropDownClosed
        findPeopleList(String.Empty)
    End Sub

    Private Function findPeopleList(ByVal namesSQL As String) As Boolean
        Dim personDAL As TKPC.DAL.PersonDAL = TKPC.DAL.PersonDAL.getInstance()
        Dim ds As DataSet = Nothing
        Dim statusOptions As String = ""
        Dim titleYear As String = tscbYear.Text
        findPeopleList = False

        Try
            '0 - 전체 데이터
            If chkStatus1.Checked = True And chkStatus2.Checked = True And chkStatus3.Checked = True And chkStatus4.Checked = True Then
                ds = personDAL.findPeopleList(cbFindWhat.SelectedIndex, txtFind.Text, 0, cbType.SelectedIndex, namesSQL, titleYear, txtAgeFrom.Text, txtAgeTo.Text)
                statusOptions = "전체"

                '1 - 등록교인만
            ElseIf chkStatus1.Checked = True And chkStatus2.Checked = False And chkStatus3.Checked = False And chkStatus4.Checked = False Then
                ds = personDAL.findPeopleList(cbFindWhat.SelectedIndex, txtFind.Text, 1, cbType.SelectedIndex, namesSQL, titleYear, txtAgeFrom.Text, txtAgeTo.Text)
                statusOptions = "등록교인"

                '2 - 등록교인 + 새교우
            ElseIf chkStatus1.Checked = True And chkStatus2.Checked = True And chkStatus3.Checked = False And chkStatus4.Checked = False Then
                ds = personDAL.findPeopleList(cbFindWhat.SelectedIndex, txtFind.Text, 2, cbType.SelectedIndex, namesSQL, titleYear, txtAgeFrom.Text, txtAgeTo.Text)
                statusOptions = "등록교인,새교우"

                '3 - 등록교인 + 새교우 + 이적교인
            ElseIf chkStatus1.Checked = True And chkStatus2.Checked = True And chkStatus3.Checked = True And chkStatus4.Checked = False Then
                ds = personDAL.findPeopleList(cbFindWhat.SelectedIndex, txtFind.Text, 3, cbType.SelectedIndex, namesSQL, titleYear, txtAgeFrom.Text, txtAgeTo.Text)
                statusOptions = "등록교인,새교우,이적교인"

                '4 - 등록교인 + 새교우 + 고인
            ElseIf chkStatus1.Checked = True And chkStatus2.Checked = True And chkStatus3.Checked = False And chkStatus4.Checked = True Then
                ds = personDAL.findPeopleList(cbFindWhat.SelectedIndex, txtFind.Text, 4, cbType.SelectedIndex, namesSQL, titleYear, txtAgeFrom.Text, txtAgeTo.Text)
                statusOptions = "등록교인,새교우,고인"

                '5 - 등록교인 + 이적교인
            ElseIf chkStatus1.Checked = True And chkStatus2.Checked = False And chkStatus3.Checked = True And chkStatus4.Checked = False Then
                ds = personDAL.findPeopleList(cbFindWhat.SelectedIndex, txtFind.Text, 5, cbType.SelectedIndex, namesSQL, titleYear, txtAgeFrom.Text, txtAgeTo.Text)
                statusOptions = "등로교인,이적교인"

                '6 - 등록교인 + 이적교인 + 고인
            ElseIf chkStatus1.Checked = True And chkStatus2.Checked = False And chkStatus3.Checked = True And chkStatus4.Checked = True Then
                ds = personDAL.findPeopleList(cbFindWhat.SelectedIndex, txtFind.Text, 6, cbType.SelectedIndex, namesSQL, titleYear, txtAgeFrom.Text, txtAgeTo.Text)
                statusOptions = "등록교인,이적교인,고인"

                '7 - 등록교인 + 고인
            ElseIf chkStatus1.Checked = True And chkStatus2.Checked = False And chkStatus3.Checked = False And chkStatus4.Checked = True Then
                ds = personDAL.findPeopleList(cbFindWhat.SelectedIndex, txtFind.Text, 7, cbType.SelectedIndex, namesSQL, titleYear, txtAgeFrom.Text, txtAgeTo.Text)
                statusOptions = "등록교인,고인"

                '8 - 새교우만
            ElseIf chkStatus1.Checked = False And chkStatus2.Checked = True And chkStatus3.Checked = False And chkStatus4.Checked = False Then
                ds = personDAL.findPeopleList(cbFindWhat.SelectedIndex, txtFind.Text, 8, cbType.SelectedIndex, namesSQL, titleYear, txtAgeFrom.Text, txtAgeTo.Text)
                statusOptions = "새교우"

                '9 - 새교우 + 이적교인
            ElseIf chkStatus1.Checked = False And chkStatus2.Checked = True And chkStatus3.Checked = True And chkStatus4.Checked = False Then
                ds = personDAL.findPeopleList(cbFindWhat.SelectedIndex, txtFind.Text, 9, cbType.SelectedIndex, namesSQL, titleYear, txtAgeFrom.Text, txtAgeTo.Text)
                statusOptions = "새교우,이적교인"

                '10 - 새교우 + 이적교인 + 고인
            ElseIf chkStatus1.Checked = False And chkStatus2.Checked = True And chkStatus3.Checked = True And chkStatus4.Checked = True Then
                ds = personDAL.findPeopleList(cbFindWhat.SelectedIndex, txtFind.Text, 10, cbType.SelectedIndex, namesSQL, titleYear, txtAgeFrom.Text, txtAgeTo.Text)
                statusOptions = "새교우,이적교인,고인"

                '11 - 새교우 + 고인
            ElseIf chkStatus1.Checked = False And chkStatus2.Checked = True And chkStatus3.Checked = False And chkStatus4.Checked = True Then
                ds = personDAL.findPeopleList(cbFindWhat.SelectedIndex, txtFind.Text, 11, cbType.SelectedIndex, namesSQL, titleYear, txtAgeFrom.Text, txtAgeTo.Text)
                statusOptions = "새교우,고인"

                '12 - 이적교인만
            ElseIf chkStatus1.Checked = False And chkStatus2.Checked = False And chkStatus3.Checked = True And chkStatus4.Checked = False Then
                ds = personDAL.findPeopleList(cbFindWhat.SelectedIndex, txtFind.Text, 12, cbType.SelectedIndex, namesSQL, titleYear, txtAgeFrom.Text, txtAgeTo.Text)
                statusOptions = "이적교인"

                '13 - 이적교인 + 고인
            ElseIf chkStatus1.Checked = False And chkStatus2.Checked = False And chkStatus3.Checked = True And chkStatus4.Checked = True Then
                ds = personDAL.findPeopleList(cbFindWhat.SelectedIndex, txtFind.Text, 13, cbType.SelectedIndex, namesSQL, titleYear, txtAgeFrom.Text, txtAgeTo.Text)
                statusOptions = "이적교인,고인"

                '14 - 고인만
            ElseIf chkStatus1.Checked = False And chkStatus2.Checked = False And chkStatus3.Checked = False And chkStatus4.Checked = True Then
                ds = personDAL.findPeopleList(cbFindWhat.SelectedIndex, txtFind.Text, 14, cbType.SelectedIndex, namesSQL, titleYear, txtAgeFrom.Text, txtAgeTo.Text)
                statusOptions = "고인"

                '15 - 모두 아님
            ElseIf chkStatus1.Checked = False And chkStatus2.Checked = False And chkStatus3.Checked = False And chkStatus4.Checked = False Then
                ds = personDAL.findPeopleList(cbFindWhat.SelectedIndex, txtFind.Text, 15, cbType.SelectedIndex, namesSQL, titleYear, txtAgeFrom.Text, txtAgeTo.Text)
                statusOptions = "무적교인"

            End If
            m_Global.setNavigatorBinding(ds, 0, False, "교우리스트, 등록옵션 - " + statusOptions)

        Catch ex As Exception
            Logger.LogError(ex.ToString)
            pnFind.Visible = False
            findPeopleList = True
        Finally
            personDAL = Nothing
            ds = Nothing
        End Try
        Return findPeopleList
    End Function

    Private Sub txtAgeFrom_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtAgeFrom.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtAgeTo_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtAgeTo.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub 암호변경ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 암호변경ToolStripMenuItem.Click
        Me.Cursor = Cursors.WaitCursor
        frmResetPassword.ShowDialog()
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub cbPastor_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbPastor.SelectedIndexChanged
        '0-주단위 헌금 집계 현황
        '1-주단위 헌금 통계
        '2-월단위 헌금 집계 현황
        '3-월단위 헌금 통계
        '4-년단위 헌금 집계 현황
        '5-년단위 월별 헌금
        '6-년단위 헌금 통계
        '7-부부 리스트(남편이 먼저)
        '8-부부 리스트(부인이 먼저)
        '9-가족 리스트
        '10-고인 리스트
        '11-유아세례 리스트
        '12-입교 리스트
        '13-세례 리스트
        '14-새교우 리스트
        '15-등록교인 리스트
        '16-이적교인 리스트
        '17-생일 리스트
        '18-교적카드
        '19-헌금기록이 없는 교우 리스트
        '20-교인 주소록 설문 양식

        If cbPastor.SelectedIndex = 0 Then '주단위 헌금
            cbCurrentEnableWhich(False, True)
            enableTypeCheckBoxPastor(cbPastor.SelectedIndex)
            txtFind.SelectAll()
            txtFind.Focus()
            dtpCurrentStartDate.Visible = True
            dtpCurrentStartDate.Format = DateTimePickerFormat.Long
            dtpCurrentEndDate.Format = DateTimePickerFormat.Long
            changeStatusCheckBox(True, True, False, False, True, True, True, True)
            familyMode = False

        ElseIf cbPastor.SelectedIndex = 1 Then '주단위 헌금통계
            cbCurrentEnableWhich(False, True)
            enableTypeCheckBoxPastor(cbPastor.SelectedIndex)
            dtpCurrentStartDate.Visible = True
            dtpCurrentStartDate.Format = DateTimePickerFormat.Long
            dtpCurrentEndDate.Format = DateTimePickerFormat.Long
            changeStatusCheckBox(True, True, False, False, True, True, True, True)
            familyMode = False

        ElseIf cbPastor.SelectedIndex = 2 Then '월단위 헌금
            cbCurrentEnableWhich(False, True)
            enableTypeCheckBoxPastor(cbPastor.SelectedIndex)
            txtFind.SelectAll()
            txtFind.Focus()
            dtpCurrentStartDate.Visible = True
            dtpCurrentStartDate.Format = DateTimePickerFormat.Custom
            dtpCurrentStartDate.CustomFormat = "yyyy-MM"
            dtpCurrentEndDate.Format = DateTimePickerFormat.Custom
            dtpCurrentEndDate.CustomFormat = "yyyy-MM"
            changeStatusCheckBox(True, True, False, False, True, True, True, True)
            familyMode = False

        ElseIf cbPastor.SelectedIndex = 3 Then '월단위 헌금통계
            cbCurrentEnableWhich(False, True)
            enableTypeCheckBoxPastor(cbPastor.SelectedIndex)
            dtpCurrentStartDate.Visible = True
            dtpCurrentStartDate.Format = DateTimePickerFormat.Custom
            dtpCurrentStartDate.CustomFormat = "yyyy-MM"
            dtpCurrentEndDate.Format = DateTimePickerFormat.Custom
            dtpCurrentEndDate.CustomFormat = "yyyy-MM"
            changeStatusCheckBox(True, True, False, False, True, True, True, True)
            familyMode = False

        ElseIf cbPastor.SelectedIndex = 4 Or '년단위 헌금
            cbCurrent.SelectedIndex = 5 Then  '년단위 월별 헌금
            cbCurrentEnableWhich(False, True)
            enableTypeCheckBoxPastor(cbPastor.SelectedIndex)
            txtFind.SelectAll()
            txtFind.Focus()
            dtpCurrentStartDate.Visible = True
            dtpCurrentStartDate.Format = DateTimePickerFormat.Custom
            dtpCurrentStartDate.CustomFormat = "yyyy"
            dtpCurrentEndDate.Format = DateTimePickerFormat.Custom
            dtpCurrentEndDate.CustomFormat = "yyyy"
            changeStatusCheckBox(True, True, False, False, True, True, True, True)
            familyMode = False

        ElseIf cbPastor.SelectedIndex = 6 Then '년단위 헌금통계
            cbCurrentEnableWhich(False, True)
            enableTypeCheckBoxPastor(cbPastor.SelectedIndex)
            dtpCurrentStartDate.Visible = True
            dtpCurrentStartDate.Format = DateTimePickerFormat.Custom
            dtpCurrentStartDate.CustomFormat = "yyyy"
            dtpCurrentEndDate.Format = DateTimePickerFormat.Custom
            dtpCurrentEndDate.CustomFormat = "yyyy"
            changeStatusCheckBox(True, True, False, False, True, True, True, True)
            familyMode = False

        ElseIf cbPastor.SelectedIndex = 7 Or '남편이 먼저 부부리스트
               cbPastor.SelectedIndex = 8 Or '부인이 먼저 부부리스트
               cbPastor.SelectedIndex = 9 Then '가족 리스트
            dtpCurrentStartDate.Visible = True
            cbCurrentEnableWhich(True, False)
            changeStatusCheckBox(True, True, False, False, True, True, True, True)
            enableTypeCheckBoxPastor(cbPastor.SelectedIndex)
            txtFind.SelectAll()
            txtFind.Focus()

            If cbPastor.SelectedIndex = 7 Or cbPastor.SelectedIndex = 8 Then
                familyMode = False
            End If

        ElseIf cbPastor.SelectedIndex = 10 Or '사망자 리스트
            cbPastor.SelectedIndex = 11 Or '유아세례 리스트
            cbPastor.SelectedIndex = 12 Or '입교 리스트
            cbPastor.SelectedIndex = 13 Or '세례 리스트
            cbPastor.SelectedIndex = 14 Or ' 새교우 리스트
            cbPastor.SelectedIndex = 15 Or '등록 리스트
            cbPastor.SelectedIndex = 16 Then '이적 리스트
            dtpCurrentStartDate.Visible = True
            cbCurrentEnableWhich(True, True)
            enableTypeCheckBoxPastor(cbPastor.SelectedIndex)
            txtFind.SelectAll()
            txtFind.Focus()
            dtpCurrentStartDate.Format = DateTimePickerFormat.Long
            dtpCurrentEndDate.Format = DateTimePickerFormat.Long

            If cbPastor.SelectedIndex = 10 Then '고인
                changeStatusCheckBox(False, False, False, True, True, True, True, True)
            ElseIf cbPastor.SelectedIndex = 14 Then '새교우
                changeStatusCheckBox(False, True, False, False, True, True, True, True)
            ElseIf cbPastor.SelectedIndex = 16 Then '이적
                changeStatusCheckBox(False, False, True, False, True, True, True, True)
            ElseIf cbPastor.SelectedIndex = 15 Then '등록
                changeStatusCheckBox(True, False, False, False, True, True, True, True)
            Else
                changeStatusCheckBox(True, True, False, False, True, True, True, True)
            End If

        ElseIf cbPastor.SelectedIndex = 17 Then '생일리스트
            cbCurrentEnableWhich(True, True)
            enableTypeCheckBoxPastor(cbPastor.SelectedIndex)
            txtFind.SelectAll()
            txtFind.Focus()
            dtpCurrentStartDate.Visible = True
            dtpCurrentStartDate.Format = DateTimePickerFormat.Custom
            dtpCurrentStartDate.CustomFormat = "M월"
            dtpCurrentEndDate.Format = DateTimePickerFormat.Custom
            dtpCurrentEndDate.CustomFormat = "M월"
            changeStatusCheckBox(True, True, False, False, True, True, True, True)

        ElseIf cbPastor.SelectedIndex = 18 Then '교적카드
            cbCurrentEnableWhich(True, False)
            enableTypeCheckBoxPastor(cbPastor.SelectedIndex)
            dtpCurrentStartDate.Visible = True
            txtFind.SelectAll()
            txtFind.Focus()
            changeStatusCheckBox(True, True, False, False, True, True, True, True)
            familyMode = True

        ElseIf cbPastor.SelectedIndex = 19 Then '헌금기록이 없는 교우 리스트
            cbCurrentEnableWhich(True, True)
            enableTypeCheckBoxPastor(cbPastor.SelectedIndex)
            dtpCurrentStartDate.Visible = True
            dtpCurrentStartDate.Format = DateTimePickerFormat.Long
            dtpCurrentEndDate.Format = DateTimePickerFormat.Long
            Dim d3weekagoSunday As Date = Now.AddDays(-(Now.DayOfWeek) - 21)
            dtpCurrentStartDate.Value = d3weekagoSunday
            changeStatusCheckBox(True, True, False, False, True, True, True, True)
            txtFind.SelectAll()
            txtFind.Focus()
            familyMode = False

        ElseIf cbPastor.SelectedIndex = 20 Then '설문
            cbCurrentEnableWhich(True, False)
            enableTypeCheckBoxPastor(cbPastor.SelectedIndex)
            dtpCurrentStartDate.Visible = True
            txtFind.SelectAll()
            txtFind.Focus()
            changeStatusCheckBox(True, True, False, False, True, True, True, True)
        End If
    End Sub

    Private Sub btGoPastor_Click(sender As Object, e As EventArgs) Handles btGoPastor.Click
        txtFind.Clear()
        pnFind.Visible = False

        Dim title As String = ""

        If cbPastor.SelectedIndex = 0 Or cbPastor.SelectedIndex = 1 Or
           cbPastor.SelectedIndex = 2 Or cbPastor.SelectedIndex = 3 Or
           cbPastor.SelectedIndex = 4 Or cbPastor.SelectedIndex = 5 Or
           cbPastor.SelectedIndex = 6 Then '주단위, 년단위
            Dim offeringDAL As TKPC.DAL.OfferingDAL = TKPC.DAL.OfferingDAL.getInstance()
            Dim ds As DataSet = Nothing
            Dim selectedStartDate As String = dtpCurrentStartDate.Value.ToString("yyyy-MM-dd")
            Dim selectedEndDate As String = dtpCurrentEndDate.Value.ToString("yyyy-MM-dd")

            Try
                If m_Global.compare2Dates(CType(selectedStartDate, Date), CType(selectedEndDate, Date)) = False Then
                    MsgBox("시작날짜가 끝날짜보다 더 클 수 없습니다. 다시 선택해주십시오.", Nothing, "주의")
                    dtpCurrentStartDate.Focus()
                    Exit Sub
                End If

                Me.Cursor = Cursors.WaitCursor

                Dim personIds As String = ""
                Dim cp As List(Of String) = readPersonIdFromDgChosen(0)
                If cp.Count > 0 Then
                    personIds = String.Join(",", cp)
                    prevFoundPersonList = cp
                End If

                If cbPastor.SelectedIndex = 0 Then '주단위 헌금

                    If cp.Count > 0 Then
                        If MsgBox("개인의 헌금내역을 볼 수 없습니다. " + vbCrLf + "옮겨담은 개인 데이터를 모두 삭제합니다. 진행하시겠습니까? ", MsgBoxStyle.YesNo, "선택") = vbYes Then
                            dgChosen.Rows.Clear()
                            cp.Clear()
                            personIds = ""
                            findPeopleHashTable.Clear()
                        Else
                            Me.Cursor = Cursors.Default
                            Exit Sub
                        End If
                    End If

                    If chkType.Checked = False And chkAll.Checked = False And dgChosen.RowCount = 0 Then
                        chkType.Checked = True 'default
                    End If

                    Dim datesList As List(Of String) = m_Global.getOfferingDates(0, selectedStartDate, selectedEndDate)
                    ds = offeringDAL.getOfferingFlexibleReport1(chkType.Checked, selectedStartDate, selectedEndDate, personIds, chkAll.Checked, datesList)

                ElseIf cbPastor.SelectedIndex = 1 Then '주단위 헌금통계
                    ds = offeringDAL.getOfferingFlexibleReport2(selectedStartDate, selectedEndDate)

                ElseIf cbPastor.SelectedIndex = 2 Then '월단위 헌금
                    If chkType.Checked = False And chkAll.Checked = False And dgChosen.RowCount = 0 Then
                        chkType.Checked = True 'default
                    End If

                    If cp.Count > 0 Then
                        If MsgBox("개인의 헌금내역을 볼 수 없습니다. " + vbCrLf + "옮겨담은 개인 데이터를 모두 삭제합니다. 진행하시겠습니까? ", MsgBoxStyle.YesNo, "선택") = vbYes Then
                            dgChosen.Rows.Clear()
                            cp.Clear()
                            personIds = ""
                            findPeopleHashTable.Clear()
                        Else
                            Me.Cursor = Cursors.Default
                            Exit Sub
                        End If
                    End If

                    selectedStartDate = m_Global.formatDate(1, selectedStartDate, "01")
                    selectedEndDate = m_Global.formatDate(1, selectedEndDate, "31")

                    ds = offeringDAL.getOfferingFlexibleReport3(chkType.Checked, selectedStartDate, selectedEndDate, personIds, chkAll.Checked)

                ElseIf cbPastor.SelectedIndex = 3 Then '월단위 헌금통계
                    selectedStartDate = m_Global.formatDate(1, selectedStartDate, "01")
                    selectedEndDate = m_Global.formatDate(1, selectedEndDate, "31")

                    ds = offeringDAL.getOfferingFlexibleReport4(selectedStartDate, selectedEndDate)

                ElseIf cbPastor.SelectedIndex = 4 Then '년단위 헌금
                    If chkType.Checked = False And chkAll.Checked = False And dgChosen.RowCount = 0 Then
                        chkType.Checked = True 'default
                    End If

                    If cp.Count > 0 Then
                        If MsgBox("개인의 헌금내역을 볼 수 없습니다. " + vbCrLf + "옮겨담은 개인 데이터를 모두 삭제합니다. 진행하시겠습니까? ", MsgBoxStyle.YesNo, "선택") = vbYes Then
                            dgChosen.Rows.Clear()
                            cp.Clear()
                            personIds = ""
                            findPeopleHashTable.Clear()
                        Else
                            Me.Cursor = Cursors.Default
                            Exit Sub
                        End If
                    End If

                    selectedStartDate = dtpCurrentStartDate.Value.ToString("yyyy")
                    selectedEndDate = dtpCurrentEndDate.Value.ToString("yyyy")

                    ds = offeringDAL.getOfferingFlexibleReport5(chkType.Checked, selectedStartDate, selectedEndDate, personIds, chkAll.Checked)

                ElseIf cbPastor.SelectedIndex = 5 Then '년단위 월별 헌금

                    If cp.Count > 0 Then
                        If MsgBox("개인의 헌금내역을 볼 수 없습니다. " + vbCrLf + "옮겨담은 개인 데이터를 모두 삭제합니다. 진행하시겠습니까? ", MsgBoxStyle.YesNo, "선택") = vbYes Then
                            dgChosen.Rows.Clear()
                            cp.Clear()
                            personIds = ""
                            findPeopleHashTable.Clear()
                        Else
                            Me.Cursor = Cursors.Default
                            Exit Sub
                        End If
                    End If

                    selectedStartDate = dtpCurrentStartDate.Value.ToString("yyyy")
                    selectedEndDate = dtpCurrentEndDate.Value.ToString("yyyy")
                    ds = offeringDAL.getOfferingFlexibleReport6(chkType.Checked, selectedStartDate, selectedEndDate, personIds)

                ElseIf cbPastor.SelectedIndex = 6 Then
                    ds = offeringDAL.getOfferingFlexibleReport7(selectedStartDate, selectedEndDate)
                End If

                title = cbPastor.Text + ", 기간:  " + selectedStartDate + " ~ " + selectedEndDate

                m_Global.setNavigatorBinding(ds, 1, False, title)

            Catch ex As Exception
                Logger.LogError(ex.ToString)
            Finally
                Me.Cursor = Cursors.Default
                offeringDAL = Nothing
                ds = Nothing
            End Try

        ElseIf cbPastor.SelectedIndex = 7 Or cbPastor.SelectedIndex = 8 Then '부부리스트
            Me.Cursor = Cursors.WaitCursor
            Dim personDAL As TKPC.DAL.PersonDAL = TKPC.DAL.PersonDAL.getInstance()
            Dim ds As DataSet = Nothing

            Try
                Dim personNames As String = ""
                Dim cp As List(Of String) = readPersonIdFromDgChosen(1)
                If cp.Count > 0 Then
                    personNames = String.Join("','", cp)
                    personNames = "'" + personNames + "'"
                    prevFoundPersonList = cp
                End If

                Dim statusOptions As String = ""
                '0 - 전체 데이터
                If chkStatus1.Checked = True And chkStatus2.Checked = True And chkStatus3.Checked = True And chkStatus4.Checked = True Then
                    ds = personDAL.getCoupleList(cbPastor.SelectedIndex - 7, personNames, 0)
                    statusOptions = "전체"

                    '1 - 등록교인만
                ElseIf chkStatus1.Checked = True And chkStatus2.Checked = False And chkStatus3.Checked = False And chkStatus4.Checked = False Then
                    ds = personDAL.getCoupleList(cbPastor.SelectedIndex - 7, personNames, 1)
                    statusOptions = "등록교인"

                    '2 - 등록교인 + 새교우
                ElseIf chkStatus1.Checked = True And chkStatus2.Checked = True And chkStatus3.Checked = False And chkStatus4.Checked = False Then
                    ds = personDAL.getCoupleList(cbPastor.SelectedIndex - 7, personNames, 2)
                    statusOptions = "등록교인,새교우"

                    '3 - 등록교인 + 새교우 + 이적교인
                ElseIf chkStatus1.Checked = True And chkStatus2.Checked = True And chkStatus3.Checked = True And chkStatus4.Checked = False Then
                    ds = personDAL.getCoupleList(cbPastor.SelectedIndex - 7, personNames, 3)
                    statusOptions = "등록교인,새교우,이적교인"

                    '4 - 등록교인 + 새교우 + 고인
                ElseIf chkStatus1.Checked = True And chkStatus2.Checked = True And chkStatus3.Checked = False And chkStatus4.Checked = True Then
                    ds = personDAL.getCoupleList(cbPastor.SelectedIndex - 7, personNames, 4)
                    statusOptions = "등록교인,새교우,고인"

                    '5 - 등록교인 + 이적교인
                ElseIf chkStatus1.Checked = True And chkStatus2.Checked = False And chkStatus3.Checked = True And chkStatus4.Checked = False Then
                    ds = personDAL.getCoupleList(cbPastor.SelectedIndex - 7, personNames, 5)
                    statusOptions = "등록교인,이적교인"

                    '6 - 등록교인 + 이적교인 + 고인
                ElseIf chkStatus1.Checked = True And chkStatus2.Checked = False And chkStatus3.Checked = True And chkStatus4.Checked = True Then
                    ds = personDAL.getCoupleList(cbPastor.SelectedIndex - 7, personNames, 6)
                    statusOptions = "등록교인,이적교인,고인"

                    '7 - 등록교인 + 고인
                ElseIf chkStatus1.Checked = True And chkStatus2.Checked = False And chkStatus3.Checked = False And chkStatus4.Checked = True Then
                    ds = personDAL.getCoupleList(cbPastor.SelectedIndex - 7, personNames, 7)
                    statusOptions = "등록교인,고인"

                    '8 - 새교우만
                ElseIf chkStatus1.Checked = False And chkStatus2.Checked = True And chkStatus3.Checked = False And chkStatus4.Checked = False Then
                    ds = personDAL.getCoupleList(cbPastor.SelectedIndex - 7, personNames, 8)
                    statusOptions = "새교우"

                    '9 - 새교우 + 이적교인
                ElseIf chkStatus1.Checked = False And chkStatus2.Checked = True And chkStatus3.Checked = True And chkStatus4.Checked = False Then
                    ds = personDAL.getCoupleList(cbPastor.SelectedIndex - 7, personNames, 9)
                    statusOptions = "새교우,이적교인"

                    '10 - 새교우 + 이적교인 + 고인
                ElseIf chkStatus1.Checked = False And chkStatus2.Checked = True And chkStatus3.Checked = True And chkStatus4.Checked = True Then
                    ds = personDAL.getCoupleList(cbPastor.SelectedIndex - 7, personNames, 10)
                    statusOptions = "새교우,이적교인,고인"

                    '11 - 새교우 + 고인
                ElseIf chkStatus1.Checked = False And chkStatus2.Checked = True And chkStatus3.Checked = False And chkStatus4.Checked = True Then
                    ds = personDAL.getCoupleList(cbPastor.SelectedIndex - 7, personNames, 11)
                    statusOptions = "새교우,고인"

                    '12 - 이적교인만
                ElseIf chkStatus1.Checked = False And chkStatus2.Checked = False And chkStatus3.Checked = True And chkStatus4.Checked = False Then
                    ds = personDAL.getCoupleList(cbPastor.SelectedIndex - 7, personNames, 12)
                    statusOptions = "이적교인"

                    '13 - 이적교인 + 고인
                ElseIf chkStatus1.Checked = False And chkStatus2.Checked = False And chkStatus3.Checked = True And chkStatus4.Checked = True Then
                    ds = personDAL.getCoupleList(cbPastor.SelectedIndex - 7, personNames, 13)
                    statusOptions = "이적교인,고인"

                    '14 - 고인만
                ElseIf chkStatus1.Checked = False And chkStatus2.Checked = False And chkStatus3.Checked = False And chkStatus4.Checked = True Then
                    ds = personDAL.getCoupleList(cbPastor.SelectedIndex - 7, personNames, 14)
                    statusOptions = "고인"

                    '15 - 모두 아님
                ElseIf chkStatus1.Checked = False And chkStatus2.Checked = False And chkStatus3.Checked = False And chkStatus4.Checked = False Then
                    chkStatus1.Checked = True
                    ds = personDAL.getCoupleList(cbPastor.SelectedIndex - 7, personNames, 15)
                    statusOptions = "등록교인"
                End If

                title = cbPastor.Text + ", 등록옵션: " + statusOptions

                m_Global.setNavigatorBinding(ds, 1, False, title)
            Catch ex As Exception
                Logger.LogError(ex.ToString)
            Finally
                personDAL = Nothing
                ds = Nothing
                Me.Cursor = Cursors.Default
            End Try

        ElseIf cbPastor.SelectedIndex = 9 Then '가족 리스트
            familyMode = True
            prevFoundPersonList = readPersonIdFromDgChosen(0)
            getFamilyList(0, readPersonIdFromDgChosen(0))

        ElseIf cbPastor.SelectedIndex = 10 Or   '고인 리스트
               cbPastor.SelectedIndex = 11 Or   '유아세례 리스트
               cbPastor.SelectedIndex = 12 Or   '입교 리스트
               cbPastor.SelectedIndex = 13 Or   '세례 리스트
               cbPastor.SelectedIndex = 14 Or   '새교우 리스트
               cbPastor.SelectedIndex = 15 Or   '등록 리스트
               cbPastor.SelectedIndex = 16 Or   '이적 리스트
               cbPastor.SelectedIndex = 17 Then '생일 리스트

            Dim personDAL As TKPC.DAL.PersonDAL = TKPC.DAL.PersonDAL.getInstance()
            Dim ds As DataSet = Nothing
            Dim startDate As String = dtpCurrentStartDate.Text
            Dim endDate As String = dtpCurrentEndDate.Text

            If cbPastor.SelectedIndex = 17 Then '생일 리스트
                Dim startMonth As Integer = CType(startDate.Substring(0, startDate.Length - 1), Integer)
                Dim endMonth As Integer = CType(endDate.Substring(0, endDate.Length - 1), Integer)

                startDate = CType(startDate.Substring(0, startDate.Length - 1), Integer).ToString("D2")
                endDate = CType(endDate.Substring(0, endDate.Length - 1), Integer).ToString("D2")

                If m_Global.compare(startMonth, endMonth) = False Then
                    MsgBox("시작날짜가 끝날짜보다 더 클 수 없습니다. 다시 선택해주십시오.", Nothing, "주의")
                    dtpCurrentStartDate.Focus()
                    Me.Cursor = Cursors.Default
                    Exit Sub
                End If
            Else
                startDate = dtpCurrentStartDate.Value.ToString("yyyy-MM-dd")
                endDate = dtpCurrentEndDate.Value.ToString("yyyy-MM-dd")

                If m_Global.compare2Dates(CType(startDate, Date), CType(endDate, Date)) = False Then
                    MsgBox("시작날짜가 끝날짜보다 더 클 수 없습니다. 다시 선택해주십시오.", Nothing, "주의")
                    dtpCurrentStartDate.Focus()
                    Me.Cursor = Cursors.Default
                    Exit Sub
                End If
            End If

            Try
                Me.Cursor = Cursors.WaitCursor

                Dim personIds As String = ""
                Dim cp As List(Of String) = readPersonIdFromDgChosen(0)
                If cp.Count > 0 Then
                    personIds = String.Join(",", cp)
                    prevFoundPersonList = cp
                End If

                Dim statusOptions As String = ""
                '0 - 전체 데이터
                If chkStatus1.Checked = True And chkStatus2.Checked = True And chkStatus3.Checked = True And chkStatus4.Checked = True Then
                    ds = personDAL.getPeopleList(cbPastor.SelectedIndex - 10, personIds, startDate, endDate, 0)
                    statusOptions = "전체"

                    '1 - 등록교인만
                ElseIf chkStatus1.Checked = True And chkStatus2.Checked = False And chkStatus3.Checked = False And chkStatus4.Checked = False Then
                    ds = personDAL.getPeopleList(cbPastor.SelectedIndex - 10, personIds, startDate, endDate, 1)
                    statusOptions = "등록교인"

                    '2 - 등록교인 + 새교우
                ElseIf chkStatus1.Checked = True And chkStatus2.Checked = True And chkStatus3.Checked = False And chkStatus4.Checked = False Then
                    ds = personDAL.getPeopleList(cbPastor.SelectedIndex - 10, personIds, startDate, endDate, 2)
                    statusOptions = "등록교인,새교우"

                    '3 - 등록교인 + 새교우 + 이적교인
                ElseIf chkStatus1.Checked = True And chkStatus2.Checked = True And chkStatus3.Checked = True And chkStatus4.Checked = False Then
                    ds = personDAL.getPeopleList(cbPastor.SelectedIndex - 10, personIds, startDate, endDate, 3)
                    statusOptions = "등록교인,새교우,이적교인"

                    '4 - 등록교인 + 새교우 + 고인
                ElseIf chkStatus1.Checked = True And chkStatus2.Checked = True And chkStatus3.Checked = False And chkStatus4.Checked = True Then
                    ds = personDAL.getPeopleList(cbPastor.SelectedIndex - 10, personIds, startDate, endDate, 4)
                    statusOptions = "등록교인,새교우,고인"

                    '5 - 등록교인 + 이적교인
                ElseIf chkStatus1.Checked = True And chkStatus2.Checked = False And chkStatus3.Checked = True And chkStatus4.Checked = False Then
                    ds = personDAL.getPeopleList(cbPastor.SelectedIndex - 10, personIds, startDate, endDate, 5)
                    statusOptions = "등록교인,이적교인"

                    '6 - 등록교인 + 이적교인 + 고인
                ElseIf chkStatus1.Checked = True And chkStatus2.Checked = False And chkStatus3.Checked = True And chkStatus4.Checked = True Then
                    ds = personDAL.getPeopleList(cbPastor.SelectedIndex - 10, personIds, startDate, endDate, 6)
                    statusOptions = "등록교인,이적교인,고인"

                    '7 - 등록교인 + 고인
                ElseIf chkStatus1.Checked = True And chkStatus2.Checked = False And chkStatus3.Checked = False And chkStatus4.Checked = True Then
                    ds = personDAL.getPeopleList(cbPastor.SelectedIndex - 10, personIds, startDate, endDate, 7)
                    statusOptions = "등록교인,고인"

                    '8 - 새교우만
                ElseIf chkStatus1.Checked = False And chkStatus2.Checked = True And chkStatus3.Checked = False And chkStatus4.Checked = False Then
                    ds = personDAL.getPeopleList(cbPastor.SelectedIndex - 10, personIds, startDate, endDate, 8)
                    statusOptions = "새교우"

                    '9 - 새교우 + 이적교인
                ElseIf chkStatus1.Checked = False And chkStatus2.Checked = True And chkStatus3.Checked = True And chkStatus4.Checked = False Then
                    ds = personDAL.getPeopleList(cbPastor.SelectedIndex - 10, personIds, startDate, endDate, 9)
                    statusOptions = "새교우,이적교인"

                    '10 - 새교우 + 이적교인 + 고인
                ElseIf chkStatus1.Checked = False And chkStatus2.Checked = True And chkStatus3.Checked = True And chkStatus4.Checked = True Then
                    ds = personDAL.getPeopleList(cbPastor.SelectedIndex - 10, personIds, startDate, endDate, 10)
                    statusOptions = "새교우,이적교인,고인"

                    '11 - 새교우 + 고인
                ElseIf chkStatus1.Checked = False And chkStatus2.Checked = True And chkStatus3.Checked = False And chkStatus4.Checked = True Then
                    ds = personDAL.getPeopleList(cbPastor.SelectedIndex - 10, personIds, startDate, endDate, 11)
                    statusOptions = "새교우,고인"

                    '12 - 이적교인만
                ElseIf chkStatus1.Checked = False And chkStatus2.Checked = False And chkStatus3.Checked = True And chkStatus4.Checked = False Then
                    ds = personDAL.getPeopleList(cbPastor.SelectedIndex - 10, personIds, startDate, endDate, 12)
                    statusOptions = "이적교인"

                    '13 - 이적교인 + 고인
                ElseIf chkStatus1.Checked = False And chkStatus2.Checked = False And chkStatus3.Checked = True And chkStatus4.Checked = True Then
                    ds = personDAL.getPeopleList(cbPastor.SelectedIndex - 10, personIds, startDate, endDate, 13)
                    statusOptions = "이적교인,고인"

                    '14 - 고인만
                ElseIf chkStatus1.Checked = False And chkStatus2.Checked = False And chkStatus3.Checked = False And chkStatus4.Checked = True Then
                    ds = personDAL.getPeopleList(cbPastor.SelectedIndex - 10, personIds, startDate, endDate, 14)
                    statusOptions = "고인"

                    '15 - 모두 아님
                ElseIf chkStatus1.Checked = False And chkStatus2.Checked = False And chkStatus3.Checked = False And chkStatus4.Checked = False Then
                    chkStatus1.Checked = True
                    ds = personDAL.getPeopleList(cbPastor.SelectedIndex - 10, personIds, startDate, endDate, 15)
                    statusOptions = "등록교인"

                End If

                title = cbPastor.Text + ", 등록옵션: " + statusOptions
                m_Global.setNavigatorBinding(ds, 1, True, title)

                If chkAll.Checked = True Then '가족모드로 보기
                    familyMode = True
                    dgChosen.Rows.Clear()

                    If ugResult.Rows.Count > 0 Then
                        With ugResult
                            For i As Integer = 0 To .Rows.Count - 1
                                Dim status As String = .Rows(i).Cells("status").Text
                                Dim deceased As String = .Rows(i).Cells("deceased_date").Text
                                If String.IsNullOrEmpty(deceased) = True Then
                                    If status = "church_member" Then
                                        status = "등록"
                                    ElseIf status = "left_church" Then
                                        status = "이적"
                                    ElseIf status = "new_family" Then
                                        status = "새교우"
                                    End If
                                Else
                                    status = "고인"
                                End If

                                dgChosen.Rows.Add(CType(.Rows(i).Cells("person_id").Value, String),
                                                  .Rows(i).Cells("korean_name").Text,
                                              CType(.Rows(i).Cells("spouses").Value, String),
                                              status,
                                              .Rows(i).Cells("age").Text,
                                              .Rows(i).Cells("gender").Text)
                            Next
                        End With

                        getFamilyList(0, readPersonIdFromDgChosen(0))
                        tsbtCard.Visible = True
                        tsbtSurvey.Visible = True

                    End If
                Else
                    familyMode = False
                    tsbtSurvey.Visible = False
                    tsbtCard.Visible = False
                End If

            Catch ex As Exception
                Logger.LogError(ex.ToString)
            Finally
                personDAL = Nothing
                ds = Nothing
                Me.Cursor = Cursors.Default
            End Try

        ElseIf cbPastor.SelectedIndex = 18 Then '교적카드
            'If dgChosen.RowCount = 0 Then
            '    If MsgBox("데이터의 양이 많아?", MsgBoxStyle.YesNo, "백그라운드 처리") = MsgBoxResult.Yes Then ' If you select yes in the MsgBox then it will close the window
            '        bwTKPC.RunWorkerAsync(3000)
            '    End If
            'Else
            If dgChosen.Rows.Count = 0 Then
                MsgBox("교우를 선택해주시기 바랍니다.", Nothing, "주의")
                txtFind.Focus()
            Else
                familyMode = True
                With frmVisitInfoOption
                    .cp = readPersonIdFromDgChosen(0)
                    .ShowDialog()
                End With
            End If

            'End If
        ElseIf cbPastor.SelectedIndex = 19 Then '헌금기록이 없는 사람들
            Me.Cursor = Cursors.WaitCursor
            Dim offeringDAL As TKPC.DAL.OfferingDAL = TKPC.DAL.OfferingDAL.getInstance()
            Dim dt As DataTable = Nothing

            Try
                Dim startDate As String = dtpCurrentStartDate.Value.ToString("yyyy-MM-dd")
                Dim endDate As String = dtpCurrentEndDate.Value.ToString("yyyy-MM-dd")

                If m_Global.compare2Dates(CType(startDate, Date), CType(endDate, Date)) = False Then
                    MsgBox("시작날짜가 끝날짜보다 더 클 수 없습니다. 다시 선택해주십시오.", Nothing, "주의")
                    dtpCurrentStartDate.Focus()
                    Me.Cursor = Cursors.Default
                    Exit Sub
                End If

                Dim personIds As String = ""
                Dim cp As List(Of String) = readPersonIdFromDgChosen(0)
                If cp.Count > 0 Then
                    personIds = String.Join(",", cp)
                End If

                Dim startYear As Integer = CType(dtpCurrentStartDate.Value.ToString("yyyy"), Integer)
                Dim endYear As Integer = CType(dtpCurrentEndDate.Value.ToString("yyyy"), Integer)
                Dim paramYears As String = ""

                For i As Integer = startYear To endYear
                    If i < endYear Then
                        paramYears += CType(i, String) + ","
                    Else
                        paramYears += CType(i, String)
                    End If
                Next

                dt = offeringDAL.getPeopleWithoutOffering(startDate, endDate, personIds, paramYears)
                With ugResult
                    .DataSource = dt
                    .DataBind()
                    .Text = cbCurrent.Text + ", 기간: " + startDate + " ~ " + endDate
                End With

            Catch ex As Exception
                Logger.LogError(ex.ToString)
            Finally
                offeringDAL = Nothing
                dt = Nothing
                Me.Cursor = Cursors.Default
            End Try

            'ElseIf cbCurrent.SelectedIndex = 23 Then '교인 주소록
            '    MessageBox.Show("Coming!")
            '    'writeDirectoryBook()
        ElseIf cbPastor.SelectedIndex = 20 Then '교인 주소록 설문
            familyMode = True
            writeSurveyForm(0, -1)
            familyMode = False
            'ElseIf cbCurrent.SelectedIndex = 25 Then '명찰
            '    MessageBox.Show("Coming!")
            '    'frmNameTagOption.ShowDialog()
        End If
    End Sub

    Private Sub tsbtCard_Click(sender As Object, e As EventArgs) Handles tsbtCard.Click
        If m_Global.accessLevel = m_Constant.ACCESS_LEVEL_PASTOR Then
            With frmVisitInfoOption
                .cp = readPersonIdFromDgChosen(0)
                .ShowDialog()
            End With
        Else
            writeMembershipCard()
        End If
    End Sub

    Private Sub chkAll_CheckedChanged(sender As Object, e As EventArgs) Handles chkAll.CheckedChanged
        If chkAll.Text = m_Constant.MODE_VIEW_FAMILY And chkAll.Checked = False Then
            familyMode = False
            tsbtCard.Visible = False
            tsbtSurvey.Visible = False
        End If
    End Sub

    Private Sub ugResult_ClickCellButton(sender As Object, e As CellEventArgs) Handles ugResult.ClickCellButton
        With frmDetail
            .personId = e.Cell.Row.Cells("id").Text

            If e.Cell.Row.Cells("picture").Text = "view" Then
                .photo_file_name = m_Constant.PERSON_IMAGE_FOLDER + e.Cell.Row.Cells("PhotoName").Text
            Else
                .photo_file_name = m_Constant.PERSON_IMAGE_FOLDER + "no_image.jpg"
            End If
            .ShowDialog()
        End With
    End Sub

    Private Sub 모든데이터ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles 모든데이터ToolStripMenuItem2.Click
        With frmTitleMaker
            .whichFindPrintOption = 1
            .ShowDialog()
        End With
    End Sub

    Private Sub 선택한데이터만ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles 선택한데이터만ToolStripMenuItem2.Click
        With frmTitleMaker
            .whichFindPrintOption = 2
            .ShowDialog()
        End With
    End Sub

    Private Sub 선택하지않은데이터만ToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles 선택하지않은데이터만ToolStripMenuItem2.Click
        With frmTitleMaker
            .whichFindPrintOption = 3
            .ShowDialog()
        End With
    End Sub

    Private Sub tsbtPrint_Click(sender As Object, e As EventArgs) Handles tsbtPrint.Click
        Me.Cursor = Cursors.WaitCursor
        With ugpdResultTKPC
            .Header.TextCenter = ugResult.Text
            .Header.Appearance.FontData.Underline = Infragistics.Win.DefaultableBoolean.True
            .Header.Appearance.FontData.Bold = Infragistics.Win.DefaultableBoolean.True
            uppdResultTKPC.ShowDialog()
        End With
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub Avery5160과함께ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles tsmiSurvey1.Click
        writeSurveyForm(1, 2)
    End Sub

    Private Sub 절취선이있는세대주이름과함께ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles tsmiSurvey2.Click
        writeSurveyForm(1, 1)
    End Sub

    Private Sub 라벨없이ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles tsmiSurvey3.Click
        writeSurveyForm(1, 0)
    End Sub

    Private Sub 가족타입을Parent로변경ToolStripMenuItem_Click(sender As Object, e As EventArgs)
        Me.Cursor = Cursors.WaitCursor
        frmSwitchFamilyChild.ShowDialog()
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub tsbtReset_Click(sender As Object, e As EventArgs) Handles tsbtReset.Click
        If MsgBox("아래 리포트를 모두 지웁니다. 실행하시겠습니까", MsgBoxStyle.YesNo, "리셋") = MsgBoxResult.Yes Then ' If you select yes in the MsgBox then it will close the window
            ugResult.DeleteSelectedRows()
            ugResult.DataSource = Nothing

            tsbtExcel.Visible = False
            tsbtPrint.Visible = False
            tsbtLabel.Visible = False
            tsbtSurvey.Visible = False
            tsbtCard.Visible = False
            tsbtReset.Visible = False
        End If
    End Sub

    Private Sub 가족타입Child와Parent자리이동ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 가족타입Child와Parent자리이동ToolStripMenuItem.Click
        frmSwitchFamilyChild.ShowDialog()
    End Sub
End Class

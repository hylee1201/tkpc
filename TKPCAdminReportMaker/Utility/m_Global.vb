Imports System
Imports System.Collections
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Text
Imports Infragistics.Win

Module m_Global
    '******************  For automatic IME switch  ********************************
    Public hKrLayoutId As Long
    Public hEnLayoutId As Long
    Public whichIME As Integer = 0
    Public accessLevel As String = "pastoral_admin"  '
    Public loginUserAdminId As String = ""
    Public loginUserName As String = ""
    '**********************************************************************************
    Public Function getIPv4Address() As System.Net.IPHostEntry
        Dim strHostName As String = System.Net.Dns.GetHostName()
        getIPv4Address = Net.Dns.GetHostEntry(strHostName)

        'For Each ipheal As System.Net.IPAddress In iphe.AddressList
        '    If ipheal.AddressFamily = System.Net.Sockets.AddressFamily.InterNetwork Then
        '        getIPv4Address = ipheal.ToString()
        '    End If
        'Next
        Return getIPv4Address
    End Function

    Public Sub getFamilyList(ByVal whichWindow As Integer, ByRef cp As List(Of String))
        Dim ugResult As UltraWinGrid.UltraGrid = Nothing
        Dim tspbTKPC As System.Windows.Forms.ToolStripProgressBar = Nothing

        Dim startTime As DateTime = Convert.ToDateTime(DateTime.Now.ToLongTimeString())

        If whichWindow = 0 Then
            ugResult = mdiTKPC.ugResult
            tspbTKPC = mdiTKPC.tspbTKPC
            mdiTKPC.Cursor = Cursors.WaitCursor
        ElseIf whichWindow = 1 Then
            ugResult = frmDetail.ugResult
            tspbTKPC = frmDetail.tspbTKPC
            frmDetail.Cursor = Cursors.WaitCursor
        End If

        Dim personDAL As TKPC.DAL.PersonDAL = TKPC.DAL.PersonDAL.getInstance()
        Dim dt As DataTable = Nothing
        Dim ds As DataSet = Nothing
        Dim dtAll As DataTable = New DataTable

        Try
            If cp.Count = 0 Then
                MsgBox("데이터 옮겨담기를 하신 후에 다시 시도하시기 바랍니다.", Nothing, "주의")
                mdiTKPC.txtFind.Focus()
                Exit Sub
                'If (cp IsNot Nothing And cp.Count = 0) Or (cp Is Nothing) Then
                '    If vbYes = MessageBox.Show("등록교인들과 새교우들의 데이터만 보여집니다. " + vbCrLf + "넷트웍의 상태에 따라 최소 10분정도의 시간이 걸릴 수 있습니다. " + vbCrLf + "진행하시겠습니까?", "선택", MessageBoxButtons.YesNo) Then
                '        ds = personDAL.findPeopleList(0, Nothing, 2, 3, String.Empty, "모든년도", String.Empty, String.Empty) '등록교인들만
                '        Dim dr As DataRow
                '        For Each dr In ds.Tables(0).Rows
                '            cp.Add(dr.Field(Of String)("spouses"))
                '        Next

                '        If whichWindow = 0 Then
                '            mdiTKPC.prevFoundPersonList = cp
                '        End If
                '    Else
                '        Exit Sub
                '    End If
                'End If
            End If
            'Dim cp As List(Of String) = readPersonIdFromDgChosen(2)
            'Dim personIds As String = ""
            'personIds = String.Join(",", cp)

            'If (cp IsNot Nothing And cp.Count = 0) Or (cp Is Nothing) Then
            '    If vbYes = MessageBox.Show("데이터의 양이 많아 시간이 오래 걸릴 수 있습니다. " + vbCrLf + "진행하시겠습니까?", "선택", MessageBoxButtons.YesNo) Then
            '        bwTKPC.RunWorkerAsync()
            '    End If
            'End If

            Dim familyList As Hashtable = New Hashtable

            Dim vIdTable As Hashtable = New Hashtable
            Dim totalRecords As Integer = cp.Count - 1

            tspbTKPC.Value = 0
            tspbTKPC.Visible = True
            tspbTKPC.Maximum = totalRecords
            tspbTKPC.BackColor = Color.LightCoral

            Dim orgId As String = ""
            Dim hasSpouse As Boolean = True
            For i As Integer = 0 To totalRecords
                tspbTKPC.Value = i

                Dim theId As String = cp.Item(i).ToString
                orgId = theId
                Dim tempList As List(Of String) = Nothing
                Dim flag As Integer = 0
                Dim tempList2 As List(Of String) = Nothing
                Dim childOfSingleParent As Boolean = False

                tempList = personDAL.getPeopleFindFamilyKey(0, theId) 'check if this is a spouse
                '자식이 결혼하지 않았을 경우와 이혼한 배우자
                '자식이 결혼하지 않았을 경우에는 부모 키값을 불러올 수 있지만
                '이혼한 배우자는 부모가 없으므로 이혼한 자신의 키값을 넣어야 한다.
                '이 둘의 경우를 충족시켜야 한다.
                If (tempList Is Nothing) Or (tempList IsNot Nothing And tempList.Count = 0) Then
                    tempList = personDAL.getPeopleFindFamilyKey(1, theId) 'find other family except me (theId)
                    If (tempList IsNot Nothing And tempList.Count > 0) Then
                        Dim tempIds = String.Join(",", tempList)
                        tempList2 = personDAL.getPeopleFindFamilyKey(2, tempIds) 'check if there is a parent or child
                        If (tempList2 Is Nothing) Or (tempList2 IsNot Nothing And tempList2.Count = 0) Then
                            tempList.Add(orgId + "-ns") '배우자가 없는 부모
                        Else
                            tempList = personDAL.getPeopleFindFamilyKey(0, tempIds) 'extract only spouses
                            childOfSingleParent = True
                            orgId = ""
                        End If
                    End If
                End If

                If tempList IsNot Nothing And tempList.Count > 0 Then
                    tempList.Sort()
                    Dim vFamilyIds As String = String.Join(",", tempList)
                    If vIdTable.ContainsKey(vFamilyIds) = False Then
                        vIdTable.Add(vFamilyIds, vFamilyIds)
                    End If
                Else
                    If childOfSingleParent = False Then
                        vIdTable.Add(theId, theId)
                    End If
                End If
            Next

            If vIdTable.Count > 0 Then
                Dim vIdArray As String()
                ReDim vIdArray(vIdTable.Count - 1)
                vIdTable.Values.CopyTo(vIdArray, 0)

                familyList = New Hashtable
                Dim familyDataEnt As TKPC.Entity.FamilyDataEnt = Nothing
                For i As Integer = 0 To vIdArray.Length - 1
                    Dim theId As String = vIdArray(i)
                    If theId.Contains("-ns") Then
                        orgId = "ns"
                        theId = theId.Replace("-ns", "")
                    Else
                        orgId = ""
                    End If
                    Logger.LogInfo("theId = [" + CType(theId, String) + "]")

                    Dim familyKey As String = ""
                    dt = personDAL.getPeopleFindWithFamily(theId, orgId)
                    If dt IsNot Nothing And dt.Rows.Count > 0 Then
                        familyDataEnt = New TKPC.Entity.FamilyDataEnt
                        If dt.Rows(0) Is Nothing Then
                            familyDataEnt.koreanName = "한글이름 없음"
                            familyDataEnt.id = "0"
                        Else
                            familyDataEnt.koreanName = dt.Rows(0).Item("KoreanName").ToString
                            familyDataEnt.id = dt.Rows(0).Item("id").ToString
                        End If

                        familyDataEnt.familyDataTable = dt
                        familyKey = familyDataEnt.koreanName + "-" + dt.Rows(0).Item("id").ToString

                        If familyList.ContainsKey(familyKey) = False Then
                            familyList.Add(familyKey, familyDataEnt)
                        End If
                    Else
                        dt = personDAL.getPeopleFindWithoutFamily(theId)
                        If dt IsNot Nothing And dt.Rows.Count > 0 Then
                            familyDataEnt = New TKPC.Entity.FamilyDataEnt
                            If dt.Rows(0) Is Nothing Then
                                familyDataEnt.koreanName = "한글이름 없음"
                                familyDataEnt.id = "0"
                            Else
                                familyDataEnt.koreanName = dt.Rows(0).Item("KoreanName").ToString
                                familyDataEnt.id = dt.Rows(0).Item("id").ToString
                            End If
                            familyDataEnt.familyDataTable = dt
                            familyKey = familyDataEnt.koreanName + "-" + dt.Rows(0).Item("id").ToString

                            If familyList.ContainsKey(familyKey) = False Then
                                familyList.Add(familyKey, familyDataEnt)
                            End If
                        End If
                    End If
                Next

                Dim finalFamilyList As List(Of TKPC.Entity.FamilyDataEnt) = New List(Of TKPC.Entity.FamilyDataEnt)

                For Each familyObj As DictionaryEntry In familyList
                    finalFamilyList.Add(CType(familyObj.Value, TKPC.Entity.FamilyDataEnt))
                Next

                finalFamilyList.Sort(Function(x, y) x.koreanName.Trim.CompareTo(y.koreanName.Trim))

                For i As Integer = 0 To finalFamilyList.Count - 1
                    dtAll.Merge(finalFamilyList.Item(i).familyDataTable)
                Next
            End If

            Dim bs As BindingSource = New BindingSource
            bs.DataSource = dtAll

            With ugResult
                .DataSource = Nothing
                .DataSource = bs
                .DisplayLayout.Bands(0).Columns("PhotoName").Hidden = True
                .DisplayLayout.Bands(0).Columns("cnt").Hidden = True

                If whichWindow = 2 Then
                    .DisplayLayout.Bands(0).Columns("id").Hidden = True
                End If
                .Text = "가족리스트"
            End With

            If whichWindow = 0 Then 'Main
                If mdiTKPC.ugResult.Rows.Count = 0 Then
                    mdiTKPC.tsbtCard.Visible = False
                    mdiTKPC.tsbtSurvey.Visible = False
                    mdiTKPC.tsmiSurvey1.Visible = False
                    mdiTKPC.tsmiSurvey2.Visible = False
                    mdiTKPC.tsmiSurvey3.Visible = False
                    mdiTKPC.tsbtExcel.Visible = False
                    mdiTKPC.tsbtLabel.Visible = False
                    mdiTKPC.tsbtPrint.Visible = False
                    mdiTKPC.tsbtReset.Visible = False
                Else
                    mdiTKPC.tsbtCard.Visible = True
                    mdiTKPC.tsbtSurvey.Visible = True
                    mdiTKPC.tsmiSurvey1.Visible = True
                    mdiTKPC.tsmiSurvey2.Visible = True
                    mdiTKPC.tsmiSurvey3.Visible = True
                    mdiTKPC.tsbtExcel.Visible = True
                    mdiTKPC.tsbtLabel.Visible = False
                    mdiTKPC.tsbtPrint.Visible = True
                    mdiTKPC.tsbtReset.Visible = True
                End If
            End If

        Catch ex As Exception
            MessageBox.Show("Error at getFamilyList() : " + ex.Message)
            Logger.LogError(ex.ToString)
        Finally
            personDAL = Nothing
            dt = Nothing
            dtAll = Nothing

            If whichWindow = 0 Then
                mdiTKPC.Cursor = Cursors.Default
                mdiTKPC.tspbTKPC.Visible = False
            ElseIf whichWindow = 1 Then
                frmDetail.Cursor = Cursors.Default
                frmDetail.tspbTKPC.Visible = False
            End If

            Dim endTime As DateTime = Convert.ToDateTime(DateTime.Now.ToLongTimeString())
            Dim timeDiff = New TimeSpan(endTime.Ticks - startTime.Ticks).Seconds
            Logger.LogInfo(Convert.ToString(timeDiff))
        End Try
    End Sub

    Public Sub writeSurveyForm(ByVal whichWindow As Integer, ByVal htmlFlag As Integer, ByVal labelOption As Integer)
        Dim ugResult As UltraWinGrid.UltraGrid = Nothing
        Dim tspbTKPC As System.Windows.Forms.ToolStripProgressBar = Nothing

        If whichWindow = 0 Then
            ugResult = mdiTKPC.ugResult
            tspbTKPC = mdiTKPC.tspbTKPC
        Else
            ugResult = frmDetail.ugResult
            tspbTKPC = frmDetail.tspbTKPC
        End If

        Dim rows As Integer = ugResult.Rows.Count

        If rows > 0 Then
            Dim familyTable As SortedList = New SortedList
            Dim pairEnt As TKPC.Entity.PersonPairEnt = Nothing
            Dim familyList As List(Of TKPC.Entity.PersonPairEnt) = New List(Of TKPC.Entity.PersonPairEnt)
            Dim oddSeq As Integer = 0
            Dim familySeq As Integer = 0
            Dim headSortingKey As String = ""
            Dim skippedLevel As Integer = 0

            With ugResult
                tspbTKPC.Value = 0
                tspbTKPC.Visible = True
                tspbTKPC.Maximum = rows - 1
                tspbTKPC.BackColor = Color.LightYellow

                '.Columns(0).ColumnName = "lvl"
                '.Columns(1).ColumnName = "id"
                '.Columns(2).ColumnName = "KoreanName"
                '.Columns(3).ColumnName = "LastName"
                '.Columns(4).ColumnName = "FirstName"
                '.Columns(5).ColumnName = "Gender"
                '.Columns(6).ColumnName = "DOB"
                '.Columns(7).ColumnName = "Age"
                '.Columns(8).ColumnName = "Title"
                '.Columns(9).ColumnName = "Cell"
                '.Columns(10).ColumnName = "Home"
                '.Columns(11).ColumnName = "Work"
                '.Columns(12).ColumnName = "Email"
                '.Columns(13).ColumnName = "Type"
                '.Columns(14).ColumnName = "Address"
                '.Columns(15).ColumnName = "City"
                '.Columns(16).ColumnName = "Status"
                '.Columns(17).ColumnName = "StatusDate"
                '.Columns(18).ColumnName = "Infant_Baptism"
                '.Columns(19).ColumnName = "Infant_Baptism_Date"
                '.Columns(20).ColumnName = "Confirmation"
                '.Columns(21).ColumnName = "Confirmation_Date"
                '.Columns(22).ColumnName = "Baptism"
                '.Columns(23).ColumnName = "Baptism_Date"
                '.Columns(24).ColumnName = "PhotoName"
                '.Columns(25).ColumnName = "cnt"

                For i As Integer = 0 To rows - 1
                    tspbTKPC.Value = i

                    If .Rows(i).Cells("제외").Text = "False" Then
                        Dim lvl As String = .Rows(i).Cells("lvl").Text
                        If i <> 0 And (lvl = m_Constant.FAMILY_LEVEL_0 Or lvl = m_Constant.FAMILY_LEVEL_1) Then
                            familyList.RemoveAll(Function(pair) pair.personId1 = "")
                            If familyTable.Contains(headSortingKey) = False Then
                                familyTable.Add(headSortingKey, familyList)
                            End If
                            familyList = New List(Of TKPC.Entity.PersonPairEnt)
                            oddSeq = 0
                            familySeq = 0
                            skippedLevel = 0
                        Else
                            If skippedLevel = 1 Then
                                If i <> 0 And (lvl = m_Constant.FAMILY_LEVEL_2) Then
                                    familyList.RemoveAll(Function(pair) pair.personId1 = "")
                                    If familyTable.Contains(headSortingKey) = False Then
                                        familyTable.Add(headSortingKey, familyList)
                                    End If
                                    familyList = New List(Of TKPC.Entity.PersonPairEnt)
                                    oddSeq = 0
                                    familySeq = 0
                                    skippedLevel = 0
                                End If

                            End If
                        End If

                        Dim k As Integer = familySeq Mod 2

                        If k = 0 Then
                            pairEnt = New TKPC.Entity.PersonPairEnt
                            pairEnt.lvl1 = lvl

                            'If oddSeq = 0 Then
                            '    headSortingKey = .Rows(i).Cells(1).Text
                            'End If

                            pairEnt.personId1 = .Rows(i).Cells("id").Text
                            pairEnt.koreanName1 = .Rows(i).Cells("KoreanName").Text

                            If oddSeq = 0 Then
                                headSortingKey = pairEnt.koreanName1.Trim + "-" + CType(pairEnt.personId1, String) '.Rows(i).Cells("KoreanName").Text
                            End If

                            pairEnt.LastName1 = .Rows(i).Cells("LastName").Text
                            pairEnt.firstName1 = .Rows(i).Cells("FirstName").Text
                            pairEnt.gender1 = .Rows(i).Cells("Gender").Text
                            pairEnt.dob1 = .Rows(i).Cells("DOB").Text
                            pairEnt.age1 = .Rows(i).Cells("Age").Text
                            pairEnt.title1 = .Rows(i).Cells("Title").Text
                            pairEnt.cell1 = .Rows(i).Cells("Cell").Text
                            pairEnt.home1 = .Rows(i).Cells("Home").Text
                            pairEnt.work1 = .Rows(i).Cells("Work").Text
                            pairEnt.email1 = .Rows(i).Cells("Email").Text
                            pairEnt.type1 = m_Constant.FAMILY_TITLE + " " + CType((familySeq + 1), String) '.Rows(i).Cells("Type").Text
                            pairEnt.adddress1 = .Rows(i).Cells("Address").Text + " " + .Rows(i).Cells("City").Text
                            pairEnt.status1 = .Rows(i).Cells("Status").Text
                            pairEnt.statusDate1 = .Rows(i).Cells("StatusDate").Text
                            pairEnt.infantBaptism1 = .Rows(i).Cells("Infant_Baptism").Text
                            pairEnt.infantBaptismDate1 = .Rows(i).Cells("Infant_Baptism_Date").Text
                            pairEnt.confirmation1 = .Rows(i).Cells("Confirmation").Text
                            pairEnt.confirmationDate1 = .Rows(i).Cells("Confirmation_Date").Text
                            pairEnt.baptism1 = .Rows(i).Cells("Baptism").Text
                            pairEnt.baptismDate1 = .Rows(i).Cells("Baptism_Date").Text
                            pairEnt.photoFileName1 = .Rows(i).Cells("PhotoName").Text
                            pairEnt.totalFamily1 = CType(.Rows(i).Cells("cnt").Text, Integer)
                            familyList.Add(pairEnt)
                        Else
                            pairEnt = familyList(oddSeq)
                            pairEnt.lvl2 = lvl
                            pairEnt.personId2 = .Rows(i).Cells("id").Text
                            pairEnt.koreanName2 = .Rows(i).Cells("KoreanName").Text
                            pairEnt.LastName2 = .Rows(i).Cells("LastName").Text
                            pairEnt.firstName2 = .Rows(i).Cells("FirstName").Text
                            pairEnt.gender2 = .Rows(i).Cells("Gender").Text
                            pairEnt.dob2 = .Rows(i).Cells("DOB").Text
                            pairEnt.age2 = .Rows(i).Cells("Age").Text
                            pairEnt.title2 = .Rows(i).Cells("Title").Text
                            pairEnt.cell2 = .Rows(i).Cells("Cell").Text
                            pairEnt.home2 = .Rows(i).Cells("Home").Text
                            pairEnt.work2 = .Rows(i).Cells("Work").Text
                            pairEnt.email2 = .Rows(i).Cells("Email").Text
                            pairEnt.type2 = m_Constant.FAMILY_TITLE + " " + CType((familySeq + 1), String) '.Rows(i).Cells("Type").Text
                            pairEnt.adddress2 = .Rows(i).Cells("Address").Text + " " + .Rows(i).Cells("City").Text
                            pairEnt.status2 = .Rows(i).Cells("Status").Text
                            pairEnt.statusDate2 = .Rows(i).Cells("StatusDate").Text
                            pairEnt.infantBaptism2 = .Rows(i).Cells("Infant_Baptism").Text
                            pairEnt.infantBaptismDate2 = .Rows(i).Cells("Infant_Baptism_Date").Text
                            pairEnt.confirmation2 = .Rows(i).Cells("Confirmation").Text
                            pairEnt.confirmationDate2 = .Rows(i).Cells("Confirmation_Date").Text
                            pairEnt.baptism2 = .Rows(i).Cells("Baptism").Text
                            pairEnt.baptismDate2 = .Rows(i).Cells("Baptism_Date").Text
                            pairEnt.photoFileName2 = .Rows(i).Cells("PhotoName").Text
                            pairEnt.totalFamily2 = CType(.Rows(i).Cells("cnt").Text, Integer)
                            familyList(oddSeq) = pairEnt
                            oddSeq += 1
                        End If
                        familySeq += 1
                    Else
                        Dim lvl As String = .Rows(i).Cells("lvl").Text
                        If lvl = m_Constant.FAMILY_LEVEL_1 Then
                            skippedLevel = 1
                        End If
                    End If
                Next

                'Last family
                familyList.RemoveAll(Function(pair) pair.personId1 = "")

                If familyTable.ContainsKey(headSortingKey) = False Then
                    familyTable.Add(headSortingKey, familyList)
                End If

                If String.IsNullOrEmpty(familyTable.GetKey(0).ToString) = True Then
                    familyTable.RemoveAt(0)
                End If
            End With

            If htmlFlag >= 1 Then
                printSurveyFormHtml(familyTable)
                If whichWindow = 0 And labelOption > 0 Then
                    writeLabel(familyTable, labelOption)
                End If
            End If
            tspbTKPC.Visible = False
        End If
    End Sub
    Public Function printHeadHtml(ByVal printSize As String, ByVal isLabelPrint As Boolean, ByVal isLabelPrintOption As Integer) As String
        Dim sb As StringBuilder = New StringBuilder
        With sb
            .Append("<html moznomarginboxes mozdisallowselectionprint>" + vbCrLf)
            .Append("<title></title>" + vbCrLf)
            .Append("<head>" + vbCrLf)
            .Append("<meta charset='UTF-8'>" + vbCrLf)
            .Append("<link rel='stylesheet' href='../css/normalize.min.css'>")
            .Append("<link rel='stylesheet' href='../css/paper.css'>")

            If isLabelPrint = True Then
                .Append("<link rel='stylesheet' href='http://fonts.googleapis.com/earlyaccess/nanummyeongjo.css'> " + vbCrLf)
                .Append("<link rel='stylesheet' href='http://fonts.googleapis.com/earlyaccess/nanumpenscript.css'> " + vbCrLf)
                .Append("<link rel='stylesheet' href='http://fonts.googleapis.com/earlyaccess/jejuhallasan.css'> " + vbCrLf)
                .Append("<link rel='stylesheet' href='http://fonts.googleapis.com/earlyaccess/jejumyeongjo.css'> " + vbCrLf)
                .Append("<link rel='stylesheet' href='http://fonts.googleapis.com/earlyaccess/notosanskr.css'> " + vbCrLf)
                .Append("<link rel='stylesheet' href='http://fonts.googleapis.com/earlyaccess/nanumbrushscript.css'> " + vbCrLf)
                .Append("<link href='https://fonts.googleapis.com/css?family=Pangolin|Poppins|Candal|Hind+Guntur:700|Love+Ya+Like+A+Sister|Fontdiner+Swanky|Indie+Flower|Bad+Script' rel='stylesheet'> " + vbCrLf)
            End If

            .Append("<style type='text/css'>" + vbCrLf)
            If isLabelPrint = True Then
                'https://github.com/devp/gamecard-layout-via-html-label-templates/blob/master/template-base-avery-5160.html
                .Append("@media print { " + vbCrLf)
                .Append("   @page { " + vbCrLf)
                .Append("       size: " + printSize + ";" + vbCrLf)
                .Append("       margin: 0.14in 0.14in 0.14in 0.14in; " + vbCrLf)
                .Append("   }" + vbCrLf)
                .Append("}" + vbCrLf)
            Else
                .Append("@page { size: " + printSize + " }")
            End If

            .Append("body {" + vbCrLf)
            .Append("   background: rgb(204,204,204); " + vbCrLf)
            .Append("}" + vbCrLf)
            .Append(".alignleft {" + vbCrLf)
            .Append("   float: left;" + vbCrLf)
            .Append("   margin-left:20px;" + vbCrLf)
            .Append("}" + vbCrLf)
            .Append(".alignright {" + vbCrLf)
            .Append("   float: right;" + vbCrLf)
            .Append("   margin-right:20px;" + vbCrLf)
            .Append("}" + vbCrLf)

            If isLabelPrint = True Then
                .Append("tr { " + vbCrLf)
                .Append("   width:  100%;" + vbCrLf)
                .Append("   margin: 0;" + vbCrLf)
                .Append("}" + vbCrLf)
                .Append("td {" + vbCrLf)
                .Append("   font-size:  40px;" + vbCrLf)
                .Append("   font-weight:  bold;" + vbCrLf)
                .Append("   font-family: 'Nanum Myeongjo', serif;" + vbCrLf)
                .Append("   width: 3in;" + vbCrLf)
                .Append("   height: 1.037in;" + vbCrLf) '0.99in
                .Append("   margin: 0;" + vbCrLf)
                .Append("   vertical-align: middle;" + vbCrLf)
                .Append("   text-align: center;" + vbCrLf)
                .Append("   -moz-box-sizing: border-box; -webkit-box-sizing: border-box; box-sizing: border-box;" + vbCrLf)
                .Append("   padding: 0.05in;" + vbCrLf)

                If isLabelPrintOption = 1 Then
                    .Append("   border: 0.1px solid black;" + vbCrLf)
                End If

                .Append("   overflow:   hidden;" + vbCrLf)
                .Append("}" + vbCrLf)

                If isLabelPrintOption <> 1 Then
                    .Append("td.separator {" + vbCrLf)
                    .Append("   width:  0.12in;" + vbCrLf)
                    .Append("}" + vbCrLf)
                Else
                    .Append("td.separator {" + vbCrLf)
                    .Append("   width:  0.0in;" + vbCrLf)
                    .Append("}" + vbCrLf)
                End If
            End If

                .Append("</style>" + vbCrLf)
            .Append("<link rel ='stylesheet' href='../css/bootstrap.min.css'>" + vbCrLf)
            .Append("<script src='../js/jquery.min.js'></script>" + vbCrLf)
            .Append("<script src ='../js/bootstrap.min.js'></script>" + vbCrLf)
            .Append("</head>" + vbCrLf)
        End With
        printHeadHtml = sb.ToString
        Return printHeadHtml
    End Function
    Private Sub printSurveyFormHtml(ByRef familyTable As SortedList)
        Dim fileName As String = m_Constant.FORM_SURVEY
        Dim sb As StringBuilder = New StringBuilder
        With sb
            .Append(printHeadHtml(m_Constant.PAPER_LETTER_LANDSCAPE, False, -1))
            .Append("<body class='" + m_Constant.PAPER_LETTER_LANDSCAPE + "'>" + vbCrLf)

            For Each familyObj As DictionaryEntry In familyTable
                Dim oneFamilyList As List(Of TKPC.Entity.PersonPairEnt) = CType(familyObj.Value, List(Of TKPC.Entity.PersonPairEnt))
                'One family ArrayList
                Dim referenceId As String = "" 'Head person ID
                Dim familyCount As Integer = oneFamilyList.Count
                For i As Integer = 0 To familyCount - 1
                    Dim pair As TKPC.Entity.PersonPairEnt = oneFamilyList(i)
                    Dim lvl1 As String = pair.lvl1
                    Dim personId1 As String = pair.personId1
                    Dim koreanName1 As String = pair.koreanName1
                    Dim lastName1 As String = pair.LastName1
                    Dim firstName1 As String = pair.firstName1
                    Dim gender1 As String = pair.gender1

                    If gender1 = m_Constant.GENDER_M_EN Then
                        gender1 = m_Constant.GENDER_M_KR
                    ElseIf gender1 = m_Constant.GENDER_F_EN Then
                        gender1 = m_Constant.GENDER_F_KR
                    Else
                        gender1 = ""
                    End If

                    Dim dob1 As String = pair.dob1
                    Dim age1 As String = pair.age1
                    Dim title1 As String = pair.title1
                    Dim cell1 As String = pair.cell1
                    Dim home1 As String = pair.home1
                    Dim work1 As String = pair.work1
                    Dim email1 As String = pair.email1
                    Dim type1 As String = pair.type1
                    Dim address1 As String = pair.adddress1
                    Dim statusDate1 As String = pair.statusDate1
                    Dim infantBaptism1 As String = pair.infantBaptism1
                    Dim infantBaptismDate1 As String = pair.infantBaptismDate1
                    Dim confirmation1 As String = pair.confirmation1
                    Dim confirmationDate1 As String = pair.confirmationDate1
                    Dim baptism1 As String = pair.baptism1
                    Dim baptismDate1 As String = pair.baptismDate1
                    Dim photoFileName1 As String = pair.photoFileName1
                    Dim totalFamily1 As Integer = pair.totalFamily1

                    Dim lvl2 As String = pair.lvl2
                    Dim personId2 As String = pair.personId2
                    Dim koreanName2 As String = pair.koreanName2
                    Dim lastName2 As String = pair.LastName2
                    Dim firstName2 As String = pair.firstName2
                    Dim gender2 As String = pair.gender2

                    If gender2 = m_Constant.GENDER_M_EN Then
                        gender2 = m_Constant.GENDER_M_KR
                    ElseIf gender2 = m_Constant.GENDER_F_EN Then
                        gender2 = m_Constant.GENDER_F_KR
                    Else
                        gender2 = ""
                    End If

                    Dim dob2 As String = pair.dob2
                    Dim age2 As String = pair.age2
                    Dim title2 As String = pair.title2
                    Dim cell2 As String = pair.cell2
                    Dim home2 As String = pair.home2
                    Dim work2 As String = pair.work2
                    Dim email2 As String = pair.email2
                    Dim type2 As String = pair.type2
                    Dim address2 As String = pair.adddress2
                    Dim statusDate2 As String = pair.statusDate2
                    Dim infantBaptism2 As String = pair.infantBaptism2
                    Dim infantBaptismDate2 As String = pair.infantBaptismDate2
                    Dim confirmation2 As String = pair.confirmation2
                    Dim confirmationDate2 As String = pair.confirmationDate2
                    Dim baptism2 As String = pair.baptism2
                    Dim baptismDate2 As String = pair.baptismDate2
                    Dim photoFileName2 As String = pair.photoFileName2
                    Dim totalFamily2 As Integer = pair.totalFamily2

                    Dim k As Integer = i Mod 2
                    Dim bigTableStyle As String = "margin:0 auto; width: 95%;"
                    If k = 0 Then
                        If i = familyCount - 1 Then
                            bigTableStyle = "margin:10px; width: 49%;"
                        End If

                        .Append("<section class='sheet padding-10mm'>" + vbCrLf)
                        .Append("<table style='" + bigTableStyle + "'>" + vbCrLf)
                        .Append("<tr>" + vbCrLf)
                    End If

                    .Append("<td style='width:49%'>" + vbCrLf)
                    .Append("<table width='98%' style='margin-bottom: 3px; margin-top: 5px'>" + vbCrLf)
                    .Append("<tr width='90%'>" + vbCrLf)
                    .Append("<td style='text-align:center;font-size:18px;text-decoration: underline;'>토론토 한인장로교회 교인 신상 명세서</td>" + vbCrLf)
                    .Append("</tr>" + vbCrLf)
                    .Append("</table>" + vbCrLf)
                    .Append("<table class='table table-bordered' style='margin-bottom:0px'>" + vbCrLf)
                    .Append("<tr width='20%'>" + vbCrLf)
                    .Append("<td colspan='2' width='20%' style='text-align:center'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>" + vbCrLf)

                    If i = 0 Then
                        'If lvl1 = m_Constant.FAMILY_LEVEL_0 Or lvl1 = m_Constant.FAMILY_LEVEL_1 Then
                        '    'type1 = m_Constant.FAMILY_ME_KR
                        referenceId = personId1
                        'End If
                    End If

                    If String.IsNullOrEmpty(personId1) = True Then
                        type1 = ""
                    End If

                    'If lvl2 = m_Constant.FAMILY_LEVEL_0 Or lvl2 = m_Constant.FAMILY_LEVEL_1 Then
                    '    type2 = m_Constant.FAMILY_ME_KR
                    'End If

                    If String.IsNullOrEmpty(personId2) = True Then
                        type2 = ""
                    End If

                    Dim engName1 As String = lastName1
                    Dim engName2 As String = lastName2

                    If String.IsNullOrEmpty(firstName1) = False Then
                        engName1 = engName1 + ", " + firstName1
                    End If

                    If String.IsNullOrEmpty(firstName2) = False Then
                        engName2 = engName2 + ", " + firstName2
                    End If

                    .Append("<td colspan='2' width='40%' style='text-align:center'>" + type1 + "</td>" + vbCrLf)
                    .Append("<td colspan='2' width='40%' style='text-align:center'>" + type2 + "</td>" + vbCrLf)
                    .Append("</tr>" + vbCrLf)
                    .Append("<tr width='20%'>" + vbCrLf)
                    .Append("<td colspan='2' width='20%' style='text-align:center;font-size:14px;'>한글이름</td>" + vbCrLf)
                    .Append("<td colspan='2' width='40%' style='vertical-align:top;font-size:14px;'>" + koreanName1 + "</td>" + vbCrLf)
                    .Append("<td colspan='2' width='40%' style='vertical-align:top;font-size:14px;'>" + koreanName2 + "</td>" + vbCrLf)
                    .Append("</tr>" + vbCrLf)
                    .Append("<tr width='20%'>" + vbCrLf)
                    .Append("<td colspan='2' style='text-align:center;font-size:14px;'>공식영어명</td>" + vbCrLf)
                    .Append("<td colspan='2' style='vertical-align:top;font-size:14px;'>" + engName1 + "</td>" + vbCrLf)
                    .Append("<td colspan='2' style='vertical-align:top;font-size:14px;'>" + engName2 + "</td>" + vbCrLf)
                    .Append("</tr>" + vbCrLf)
                    .Append("<tr width='20%'>" + vbCrLf)
                    .Append("<td colspan='2' style='text-align:center;font-size:14px;'>생년월일</td>" + vbCrLf)
                    .Append("<td colspan='2' style='font-size:14px;'>" + m_Global.formatDate(0, dob1, "") + "</td>" + vbCrLf)
                    .Append("<td colspan='2' style='font-size:14px;'>" + m_Global.formatDate(0, dob2, "") + "</td>" + vbCrLf)
                    .Append("</tr>" + vbCrLf)
                    .Append("<tr width='20%'>" + vbCrLf)
                    .Append("<td colspan='2' style='text-align:center;font-size:14px;'>직분</td>" + vbCrLf)
                    .Append("<td colspan='2' style='vertical-align:top;font-size:14px;'>" + title1 + "</td>" + vbCrLf)
                    .Append("<td colspan='2' style='vertical-align:top;font-size:14px;'>" + title2 + "</td>" + vbCrLf)
                    .Append("</tr>" + vbCrLf)
                    .Append("<tr width='20%'>" + vbCrLf)
                    .Append("<td colspan='2' style='text-align:center;font-size:14px;'>성별</td>" + vbCrLf)
                    .Append("<td colspan='2' style='vertical-align:top;font-size:14px;'>" + gender1 + "</td>" + vbCrLf)
                    .Append("<td colspan='2' style='vertical-align:top;font-size:14px;'>" + gender2 + "</td>" + vbCrLf)
                    .Append("</tr>" + vbCrLf)
                    .Append("<tr width='20%'>" + vbCrLf)
                    .Append("<td colspan='2' style='text-align:center;font-size:12px;'>헌금봉투번호</td>" + vbCrLf)
                    .Append("<td colspan='2' style='vertical-align:top;font-size:14px;'></td>" + vbCrLf)
                    .Append("<td colspan='2' style='vertical-align:top;font-size:14px;'></td>" + vbCrLf)
                    .Append("</tr>" + vbCrLf)
                    .Append("<tr width='20%'>" + vbCrLf)
                    .Append("<td rowspan='3' style='text-align:center; vertical-align:middle;width:10%;font-size:14px;'>신급</td>" + vbCrLf)

                    Dim infant1html As String = "Y&nbsp;&nbsp;/&nbsp;&nbsp;N"
                    If String.IsNullOrEmpty(infantBaptism1) = False Then
                        infant1html = "Yes"
                    End If

                    Dim infantDate1html As String = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp년&nbsp;&nbsp;&nbsp월&nbsp;&nbsp;&nbsp;일"
                    If String.IsNullOrEmpty(infantBaptismDate1) = False Then
                        infantDate1html = m_Global.formatDate(0, infantBaptismDate1, "")
                    End If

                    Dim infant2html As String = "Y&nbsp;&nbsp;/&nbsp;&nbsp;N"
                    If String.IsNullOrEmpty(infantBaptism2) = False Then
                        infant2html = "Yes"
                    End If

                    Dim infantDate2html As String = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp년&nbsp;&nbsp;&nbsp월&nbsp;&nbsp;&nbsp;일"
                    If String.IsNullOrEmpty(infantBaptismDate2) = False Then
                        infantDate2html = m_Global.formatDate(0, infantBaptismDate2, "")
                    End If


                    .Append("<td style='text-align:center;font-size:14px;'>유아</td>" + vbCrLf)
                    .Append("<td style='text-align:center; width:10%;font-size:14px;'>" + infant1html + "</td>" + vbCrLf)
                    .Append("<td style='text-align:left;font-size:14px;'>" + infantDate1html + "</td>" + vbCrLf)
                    .Append("<td style='text-align:center width:10%;font-size:14px;'>" + infant2html + "</td>" + vbCrLf)
                    .Append("<td style='text-align:left;font-size:14px;'>" + infantDate2html + "</td>" + vbCrLf)
                    .Append("</tr>" + vbCrLf)
                    .Append("<tr width='20%'>" + vbCrLf)

                    Dim confirmation1html As String = "Y&nbsp;&nbsp;/&nbsp;&nbsp;N"
                    If String.IsNullOrEmpty(confirmation1) = False Then
                        confirmation1html = "Yes"
                    End If

                    Dim confirmationDate1html As String = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp년&nbsp;&nbsp;&nbsp월&nbsp;&nbsp;&nbsp;일"
                    If String.IsNullOrEmpty(confirmationDate1) = False Then
                        confirmationDate1html = m_Global.formatDate(0, confirmationDate1, "")
                    End If

                    Dim confirmation2html As String = "Y&nbsp;&nbsp;/&nbsp;&nbsp;N"
                    If String.IsNullOrEmpty(confirmation2) = False Then
                        confirmation2html = "Yes"
                    End If

                    Dim confirmationDate2html As String = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp년&nbsp;&nbsp;&nbsp월&nbsp;&nbsp;&nbsp;일"
                    If String.IsNullOrEmpty(confirmationDate2) = False Then
                        confirmationDate2html = m_Global.formatDate(0, confirmationDate2, "")
                    End If

                    .Append("<td style='text-align:center;font-size:14px;'>입교</td>" + vbCrLf)
                    .Append("<td style='text-align:center; width:10%;font-size:14px;'>" + confirmation1html + "</td>" + vbCrLf)
                    .Append("<td style='text-align:left;font-size:14px;'>" + confirmationDate1html + "</td>" + vbCrLf)
                    .Append("<td style='text-align:center; width:10%;font-size:14px;'>" + confirmation2html + "</td>" + vbCrLf)
                    .Append("<td style='text-align:left;font-size:14px;'>" + confirmationDate2html + "</td>" + vbCrLf)
                    .Append("</tr>" + vbCrLf)
                    .Append("<tr width='20%'>" + vbCrLf)
                    .Append("<td style='text-align:center;font-size:14px;'>세례</td>" + vbCrLf)

                    Dim baptism1html As String = "Y&nbsp;&nbsp;/&nbsp;&nbsp;N"
                    If String.IsNullOrEmpty(baptism1) = False Then
                        baptism1html = "Yes"
                    End If

                    Dim baptismDate1html As String = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp년&nbsp;&nbsp;&nbsp월&nbsp;&nbsp;&nbsp;일"
                    If String.IsNullOrEmpty(baptismDate1) = False Then
                        baptismDate1html = m_Global.formatDate(0, baptismDate1, "")
                    End If

                    Dim baptism2html As String = "Y&nbsp;&nbsp;/&nbsp;&nbsp;N"
                    If String.IsNullOrEmpty(baptism2) = False Then
                        baptism2html = "Yes"
                    End If

                    Dim baptismDate2html As String = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp년&nbsp;&nbsp;&nbsp월&nbsp;&nbsp;&nbsp;일"
                    If String.IsNullOrEmpty(baptismDate2) = False Then
                        baptismDate2html = m_Global.formatDate(0, baptismDate2, "")
                    End If

                    .Append("<td style='text-align:center; width:10%;font-size:14px;'>" + baptism1html + "</td>" + vbCrLf)
                    .Append("<td style='text-align:left;font-size:14px;'>" + baptismDate1html + "</td>" + vbCrLf)
                    .Append("<td style='text-align:center; width:10%;font-size:14px;'>" + baptism2html + "</td>" + vbCrLf)
                    .Append("<td style='text-align:left;font-size:14px;'>" + baptismDate2html + "</td>" + vbCrLf)
                    .Append("</tr>" + vbCrLf)
                    .Append("<tr width='20%'>" + vbCrLf)
                    .Append("<td colspan='2' style='text-align:center;font-size:14px;'>이메일주소</td>" + vbCrLf)
                    .Append("<td colspan='2' style='vertical-align:top;font-size:12px;'>" + email1 + "</td>" + vbCrLf)
                    .Append("<td colspan='2' style='vertical-align:top;font-size:12px;'>" + email2 + "</td>" + vbCrLf)
                    .Append("</tr>" + vbCrLf)
                    '.Append("<tr width='20%'>" + vbCrLf)
                    '.Append("<td colspan='2' style='text-align:center'>가족사항</td>" + vbCrLf)
                    '.Append("<td colspan='2' style='vertical-align:top'></td>" + vbCrLf)
                    '.Append("<td colspan='2' style='vertical-align:top'></td>" + vbCrLf)
                    '.Append("</tr>" + vbCrLf)
                    .Append("<tr width='20%'>" + vbCrLf)
                    .Append("<td rowspan='2' style='text-align:center; vertical-align:middle;font-size:14px;'>주소</td>" + vbCrLf)
                    .Append("<td style='text-align:center;vertical-align:middle;font-size:14px;height:65px;'>집</td>" + vbCrLf)
                    .Append("<td colspan='2' style='vertical-align:top;font-size:13px;'>" + address1 + "</td>" + vbCrLf)
                    .Append("<td colspan='2' style='vertical-align:top;font-size:13px;'>" + address2 + "</td>" + vbCrLf)
                    .Append("</tr>" + vbCrLf)
                    .Append("<tr width='20%'>" + vbCrLf)
                    .Append("<td style='text-align:center;vertical-align:middle;font-size:14px;height:65px;'>직장</td>" + vbCrLf)
                    .Append("<td colspan='2' style='vertical-align:top;height:50px;font-size:14px;'></td>" + vbCrLf)
                    .Append("<td colspan='2' style='vertical-align:top;height:50px;font-size:14px;'></td>" + vbCrLf)
                    .Append("</tr>" + vbCrLf)
                    .Append("<tr width='20%'>" + vbCrLf)
                    .Append("<td rowspan='3' style='text-align:center; vertical-align:middle;font-size:14px;'>전화<br>번호</td>" + vbCrLf)
                    .Append("<td style='text-align:center;font-size:14px;'>셀폰</td>" + vbCrLf)
                    .Append("<td colspan='2' style='vertical-align:top;font-size:14px;'>" + cell1 + "</td>" + vbCrLf)
                    .Append("<td colspan='2' style='vertical-align:top;font-size:14px;'>" + cell2 + "</td>" + vbCrLf)
                    .Append("</tr>" + vbCrLf)
                    .Append("<tr width='20%'>" + vbCrLf)
                    .Append("<td style='text-align:center;font-size:14px;'>집</td>" + vbCrLf)
                    .Append("<td colspan='2' style='vertical-align:top;font-size:14px;'>" + home1 + "</td>" + vbCrLf)
                    .Append("<td colspan='2' style='vertical-align:top;font-size:14px;'>" + home2 + "</td>" + vbCrLf)
                    .Append("</tr>" + vbCrLf)
                    .Append("<tr width='20%'>" + vbCrLf)
                    .Append("<td style='text-align:center;font-size:14px;'>직장</td>" + vbCrLf)
                    .Append("<td colspan='2' style='vertical-align:top;font-size:14px;'>" + work1 + "</td>" + vbCrLf)
                    .Append("<td colspan='2' style='vertical-align:top;font-size:14px;'>" + work2 + "</td>" + vbCrLf)
                    .Append("</tr>" + vbCrLf)
                    .Append("<tr width='20%'>" + vbCrLf)
                    .Append("<td colspan='2' style='text-align:center;font-size:14px;'>교회 등록일</td>" + vbCrLf)
                    .Append("<td colspan='2' style='font-size:14px;'>" + m_Global.formatDate(0, statusDate1, "") + "</td>" + vbCrLf)
                    .Append("<td colspan='2' style='font-size:14px;'>" + m_Global.formatDate(0, statusDate2, "") + "</td>" + vbCrLf)
                    .Append("</tr>" + vbCrLf)
                    .Append("<tr width='20%'>" + vbCrLf)
                    .Append("<td colspan='2' style='text-align:center;font-size:14px;'>소속부서</td>" + vbCrLf)
                    .Append("<td colspan='2' style='vertical-align:top;font-size:14px;'></td>" + vbCrLf)
                    .Append("<td colspan='2' style='vertical-align:top;font-size:14px;'></td>" + vbCrLf)
                    .Append("</tr>" + vbCrLf)
                    .Append("</table>" + vbCrLf)
                    .Append("</td>")

                    If k = 0 Then 'half
                        .Append("</td>" + vbCrLf)
                        .Append("<td style='width:1%'></td>" + vbCrLf)
                    End If

                    If k = 1 Or (k = 0 And i = familyCount - 1) Then
                        .Append("</tr>" + vbCrLf)
                        .Append("<tr>" + vbCrLf)
                        .Append("<td colspan='2' style='text-align:left;font-size:13px;'>* 위에 기재된 내용은 헌금영수증 발행시, 주소록 제작시 사용되는 것이니<br>영문철자를 정확히 적어주시기를 부탁드립니다.</td>" + vbCrLf)
                        .Append("</tr>" + vbCrLf)
                        .Append("<tr>" + vbCrLf)
                        .Append("<td colspan='2' style='text-align:right'>REF-" + referenceId + "</td>" + vbCrLf)
                        .Append("</tr>" + vbCrLf)
                        .Append("</table>" + vbCrLf)
                        .Append("</section>" + vbCrLf)
                    End If
                Next
            Next
            .Append("</body>" + vbCrLf)
            .Append("</html>" + vbCrLf)
        End With

        Dim fileHelper As New FileHelper
        Dim completed As Integer = fileHelper.writeToFile(fileName, sb.ToString)
        If completed = 1 Then
            System.Diagnostics.Process.Start("file:///" & m_Constant.FILE_FOLDER + fileName)
        End If
        'fileHelper.convertHtmlToPDF(m_Constant.FILE_SURVEY, sb.ToString, m_Constant.ROOT_FOLDER + "\" + fileName)
        'System.Diagnostics.Process.Start(m_Constant.ROOT_FOLDER + "\" + fileName.Replace(m_Constant.FILE_EXTENSION_HTML, m_Constant.FILE_EXTENSION_PDF))

    End Sub

    Public Sub writeMembershipCard(ByVal whichWindow As Integer, ByRef visitPersonIdTable As Hashtable, ByRef commentFlag As Boolean)
        Dim ugResult As UltraWinGrid.UltraGrid = Nothing
        Dim personDAL As TKPC.DAL.PersonDAL = TKPC.DAL.PersonDAL.getInstance()

        If whichWindow = 0 Then
            ugResult = mdiTKPC.ugResult
        Else
            ugResult = frmDetail.ugResult
        End If

        Dim rows As Integer = ugResult.Rows.Count

        If rows > 0 Then
            Dim fileName As String = "tkpc_membership_card.html"
            Dim sb As StringBuilder = New StringBuilder
            With sb
                .Append(m_Global.printHeadHtml(m_Constant.PAPER_LETTER_PORTRAIT, False, -1))
                .Append("<body class='" + m_Constant.PAPER_LETTER_PORTRAIT + "'>" + vbCrLf)

                '.Columns(0).ColumnName = "lvl"
                '.Columns(1).ColumnName = "id"
                '.Columns(2).ColumnName = "KoreanName"
                '.Columns(3).ColumnName = "LastName"
                '.Columns(4).ColumnName = "FirstName"
                '.Columns(5).ColumnName = "Gender"
                '.Columns(6).ColumnName = "DOB"
                '.Columns(7).ColumnName = "Age"
                '.Columns(8).ColumnName = "Title"
                '.Columns(9).ColumnName = "Cell"
                '.Columns(10).ColumnName = "Home"
                '.Columns(11).ColumnName = "Work"
                '.Columns(12).ColumnName = "Email"
                '.Columns(13).ColumnName = "Type"
                '.Columns(14).ColumnName = "Address"
                '.Columns(15).ColumnName = "City"
                '.Columns(16).ColumnName = "Status"
                '.Columns(17).ColumnName = "StatusDate"
                '.Columns(18).ColumnName = "Infant_Baptism"
                '.Columns(19).ColumnName = "Infant_Baptism_Date"
                '.Columns(20).ColumnName = "Confirmation"
                '.Columns(21).ColumnName = "Confirmation_Date"
                '.Columns(22).ColumnName = "Baptism"
                '.Columns(23).ColumnName = "Baptism_Date"
                '.Columns(24).ColumnName = "PhotoName"
                '.Columns(25).ColumnName = "cnt"

                Dim count As Integer = 0
                Dim personIds As String = ""
                Dim isVisitPrintQualified As Boolean = False
                For i As Integer = 0 To rows - 1
                    'z.lvl, z.pid, z.korean_name, z.last_name, z.first_name, z.gender,
                    'z.dob, z.age, z.cell, z.home, z.email, z.rtype, statusDate, cnt
                    Dim lvl As String = ugResult.Rows(i).Cells("lvl").Text
                    Dim personId As String = ugResult.Rows(i).Cells("id").Text
                    personIds += personId + ","

                    Dim koreanName As String = ugResult.Rows(i).Cells("KoreanName").Text
                    Dim lastName As String = ugResult.Rows(i).Cells("LastName").Text
                    Dim firstName As String = ugResult.Rows(i).Cells("FirstName").Text
                    Dim gender As String = ugResult.Rows(i).Cells("Gender").Text

                    If gender = m_Constant.GENDER_M_EN Then
                        gender = m_Constant.GENDER_M_KR
                    ElseIf gender = m_Constant.GENDER_F_EN Then
                        gender = m_Constant.GENDER_F_KR
                    Else
                        gender = ""
                    End If

                    Dim dob As String = ugResult.Rows(i).Cells("DOB").Text
                    Dim age As String = ugResult.Rows(i).Cells("Age").Text
                    Dim title As String = ugResult.Rows(i).Cells("Title").Text
                    Dim cell As String = ugResult.Rows(i).Cells("Cell").Text
                    Dim home As String = ugResult.Rows(i).Cells("Home").Text
                    Dim work As String = ugResult.Rows(i).Cells("Work").Text
                    Dim email As String = ugResult.Rows(i).Cells("Email").Text
                    Dim type As String = ugResult.Rows(i).Cells("Type").Text
                    Dim address As String = ugResult.Rows(i).Cells("Address").Text + " " + ugResult.Rows(i).Cells("City").Text
                    Dim status As String = ugResult.Rows(i).Cells("Status").Text
                    Dim statusDate As String = ugResult.Rows(i).Cells("StatusDate").Text
                    Dim infantBaptism = ugResult.Rows(i).Cells("Infant_Baptism").Text
                    Dim confirmation = ugResult.Rows(i).Cells("Confirmation").Text
                    Dim baptism = ugResult.Rows(i).Cells("Baptism").Text
                    Dim photoFileName = ugResult.Rows(i).Cells("PhotoName").Text
                    Dim totalFamily As Integer = CType(ugResult.Rows(i).Cells("cnt").Text, Integer)

                    If lvl = m_Constant.FAMILY_LEVEL_0 Or lvl = m_Constant.FAMILY_LEVEL_1 Then
                        count = 0
                        .Append("<section class='sheet padding-10mm'>" + vbCrLf)
                        .Append("<div id='textbox'>" + vbCrLf)
                        .Append("<p class='alignleft' style='margin-top:20px;'><img src='../template/tkpc_receipt_header.png'/></p>" + vbCrLf)
                        .Append("<p class='alignright' style='margin-top:20px'><span style='font-size:28px; font-weight:bold;'>교적카드</span></p>" + vbCrLf)
                        .Append("<div style='clear: both;'></div>" + vbCrLf)
                        .Append("</div>" + vbCrLf)
                        .Append("<table class='table table-bordered'  style='margin:0 auto; padding:15px; width: 95%'>" + vbCrLf)
                        .Append("<tr>" + vbCrLf)
                        .Append("<td style='width:64%'><span class='glyphicon glyphicon-user' aria-hidden='true'></span>&nbsp;&nbsp;" + koreanName + "</td>" + vbCrLf)

                        Dim photoHtml = "<img src='../images/person/no_image.jpg' width='270px' height='230px' style='opacity: 0.6;'/>"
                        If String.IsNullOrEmpty(photoFileName) = False Then
                            If System.IO.File.Exists(m_Constant.PERSON_IMAGE_FOLDER + photoFileName) = True Then
                                photoHtml = "<img src='../images/person/" + photoFileName + "' width='270px' height='230px'/>"
                            End If
                        End If

                        .Append("<td style='width:36%' rowspan=7'>" + photoHtml + "</td>" + vbCrLf)
                        .Append("</tr>" + vbCrLf)
                        .Append("<tr>" + vbCrLf)
                        .Append("<td><span class='glyphicon glyphicon-home' aria-hidden='true'></span>&nbsp;&nbsp;" + address + "</td>" + vbCrLf)
                        .Append("</tr>" + vbCrLf)
                        .Append("<tr>" + vbCrLf)
                        .Append("<td><span class='glyphicon glyphicon-phone' aria-hidden='true'></span>&nbsp;&nbsp;" + cell + " (셀)</td>" + vbCrLf)
                        .Append("</tr>" + vbCrLf)
                        .Append("<tr>" + vbCrLf)
                        .Append("<td><span class='glyphicon glyphicon-phone-alt' aria-hidden='true'></span>&nbsp;&nbsp;" + home + " (집) " + vbCrLf)

                        If String.IsNullOrEmpty(work) = False Then
                            .Append(" / " + work + " (직장)</td>" + vbCrLf)
                        Else
                            .Append("</td>" + vbCrLf)
                        End If

                        .Append("</tr>" + vbCrLf)
                        .Append("<tr>" + vbCrLf)
                        .Append("<td><span class='glyphicon glyphicon-heart' aria-hidden='true'></span>&nbsp;&nbsp;" + m_Global.formatDate(0, dob, "") + " (생일) " + vbCrLf)

                        If String.IsNullOrEmpty(statusDate) = False Then
                            .Append(" / " + m_Global.formatDate(0, statusDate, "") + " (등록일)</td>" + vbCrLf)
                        Else
                            .Append("</td>" + vbCrLf)
                        End If

                        .Append("</tr>" + vbCrLf)
                        .Append("<tr>" + vbCrLf)
                        .Append("<td><span class='glyphicon glyphicon-record' aria-hidden='true'></span>&nbsp;&nbsp;" + gender + vbCrLf)

                        If String.IsNullOrEmpty(baptism) = False Then
                            .Append(" / 세례</td>" + vbCrLf)
                        Else
                            If String.IsNullOrEmpty(confirmation) = False Then
                                .Append(" / 입교</td>" + vbCrLf)
                            Else
                                If String.IsNullOrEmpty(infantBaptism) = False Then
                                    .Append(" / 유아세례</td>" + vbCrLf)
                                Else
                                    .Append("</td>" + vbCrLf)
                                End If
                            End If
                        End If

                        .Append("</tr>" + vbCrLf)
                        .Append("<tr>" + vbCrLf)
                        .Append("<td><span class='glyphicon glyphicon-envelope' aria-hidden='true'></span>&nbsp;&nbsp;" + email + "</td>" + vbCrLf)
                        .Append("</tr>" + vbCrLf)
                        .Append("</table>" + vbCrLf)
                        .Append("<table class='table' style='margin:0 auto; padding:15px; width: 95%'>" + vbCrLf)
                        .Append("<tr>" + vbCrLf)
                        .Append("<td><div style='font-size:16; font-weight:bold;'>가족관계</div></td>" + vbCrLf)
                        .Append("</tr>" + vbCrLf)
                        .Append("</table>" + vbCrLf)
                        .Append("<table class='table table-bordered' style='margin:0 auto; padding:15px; width: 95%'>" + vbCrLf)
                        .Append("<thead>" + vbCrLf)
                        .Append("<tr>" + vbCrLf)
                        .Append("<th scope='col'>이름</th>" + vbCrLf)
                        .Append("<th scope='col'>관계</th>" + vbCrLf)
                        .Append("<th scope='col'>생년월일</th>" + vbCrLf)
                        .Append("<th scope='col'>나이</th>" + vbCrLf)
                        .Append("<th scope='col'>성별</th>" + vbCrLf)
                        .Append("<th scope='col'>세례유무</th>" + vbCrLf)
                        .Append("<th scope='col'>등록상태</th>" + vbCrLf)
                        .Append("</tr>" + vbCrLf)
                        .Append("</thead>" + vbCrLf)
                        .Append("<tbody>" + vbCrLf)
                    End If

                    .Append("<tr>" + vbCrLf)
                    .Append("<td style='width:20%'>" + koreanName + "</td>" + vbCrLf)

                    If count = 0 Then
                        type = m_Constant.FAMILY_ME_KR
                    End If

                    .Append("<td style='width:10%'>" + type + "</td>" + vbCrLf)
                    .Append("<td style='width:20%'>" + m_Global.formatDate(0, dob, "") + "</td>" + vbCrLf)
                    .Append("<td style='width:10%'>" + age + "</td>" + vbCrLf)
                    .Append("<td style='width:10%'>" + gender + "</td>" + vbCrLf)

                    If String.IsNullOrEmpty(baptism) = False Then
                        .Append("<td style='width:15%'>세례</td>" + vbCrLf)
                    Else
                        If String.IsNullOrEmpty(confirmation) = False Then
                            .Append("<td style='width:15%'>입교</td>" + vbCrLf)
                        Else
                            If String.IsNullOrEmpty(infantBaptism) = False Then
                                .Append("<td style='width:15%'>유아세례</td>" + vbCrLf)
                            Else
                                .Append("<td style='width:15%'></td>" + vbCrLf)
                            End If

                        End If
                    End If

                    If status = m_Constant.MEMBER_CHURCH Then
                        .Append("<td style='width:15%'>등록</td>" + vbCrLf)
                    ElseIf status = m_Constant.MEMBER_LEFT Then
                        .Append("<td style='width:15%'>이적</td>" + vbCrLf)
                    ElseIf status = m_Constant.MEMBER_NEW Then
                        .Append("<td style='width:15%'>새교우</td>" + vbCrLf)
                    Else
                        .Append("<td style='width:15%'></td>" + vbCrLf)
                    End If

                    .Append("</tr>" + vbCrLf)

                    If count = totalFamily - 1 Then
                        .Append("</tbody>" + vbCrLf)
                        .Append("</table>" + vbCrLf)

                        Dim personIdStr = personIds.Substring(0, personIds.Length - 1)

                        'Only pastor can view this information
                        If m_Global.accessLevel = m_Constant.ACCESS_LEVEL_PASTOR And visitPersonIdTable IsNot Nothing Then
                            Dim personIdsArray() As String = personIdStr.Split(New Char() {","c})
                            For j As Integer = 0 To personIdsArray.Length - 1
                                isVisitPrintQualified = visitPersonIdTable.ContainsKey(personIdsArray(j))
                                If isVisitPrintQualified = True Then
                                    Exit For
                                End If
                            Next

                            If isVisitPrintQualified Then
                                Dim visitList As List(Of TKPC.Entity.VisitEnt) = personDAL.getVisitedFamily(personIds.Substring(0, personIds.Length - 1))

                                If visitList IsNot Nothing And visitList.Count > 0 Then
                                    .Append("<table class='table' style='margin:0 auto; padding:15px; width: 95%'>" + vbCrLf)
                                    .Append("<tr>" + vbCrLf)
                                    .Append("<td><div style='font-size:16; font-weight:bold;'>마지막 심방사항</div></td>" + vbCrLf)
                                    .Append("</tr>" + vbCrLf)
                                    .Append("</table>" + vbCrLf)
                                    .Append("<table class='table table-bordered' style='margin:0 auto; padding:15px; width: 95%;'>" + vbCrLf)
                                    For k As Integer = 0 To visitList.Count - 1
                                        .Append("<tr>" + vbCrLf)
                                        .Append("<td style='width:15%;'>" + visitList(k).visitedKoreanName + "</td>" + vbCrLf)
                                        Dim visitDateHtml As String = ""
                                        If String.IsNullOrEmpty(visitList(k).visitDate) = False Then
                                            visitDateHtml = m_Global.formatDate(0, visitList(k).visitDate, "")
                                        End If
                                        .Append("<td style='width:20%;'>" + visitDateHtml + "</td>" + vbCrLf)
                                        .Append("<td style='width:65%;'>" + visitList(k).note + "</td>" + vbCrLf)
                                        .Append("</tr>" + vbCrLf)
                                    Next
                                    .Append("</table>" + vbCrLf)
                                End If
                            End If
                        End If
                    End If

                    If count = totalFamily - 1 Then
                        .Append("<table class='table' style='margin:0 auto; padding:15px; width: 95%'>" + vbCrLf)
                        .Append("<tr>" + vbCrLf)
                        .Append("<td><div style='font-size:16; font-weight:bold;'>참고사항</div></td>" + vbCrLf)
                        .Append("</tr>" + vbCrLf)

                        If commentFlag = True Then
                            Dim personIdStr = personIds.Substring(0, personIds.Length - 1)

                            Dim commentList As List(Of TKPC.Entity.CommentEnt) = personDAL.getComments(personIdStr)
                            If commentList IsNot Nothing Then
                                .Append("<tr>")
                                .Append("<td>")
                                .Append("<table class='table  table-bordered' style='margin:0 auto; padding:15px; width: 100%'>" + vbCrLf)
                                For m As Integer = 0 To commentList.Count - 1
                                    Dim commentEnt As TKPC.Entity.CommentEnt = commentList.Item(m)
                                    If commentEnt IsNot Nothing Then
                                        .Append("<tr>" + vbCrLf)
                                        .Append("<td style='width:20%;'>" + commentEnt.personName + "</td>" + vbCrLf)
                                        .Append("<td>" + commentEnt.body + "</td>" + vbCrLf)
                                        .Append("</tr>" + vbCrLf)
                                    End If
                                Next
                                .Append("</table>")
                                .Append("</td>")
                                .Append("</tr>")
                            End If
                        End If

                        .Append("</table>" + vbCrLf)
                        .Append("</section>" + vbCrLf)
                        personIds = ""
                    End If
                    count += 1
                Next
                .Append("</body>" + vbCrLf)
                .Append("</html>" + vbCrLf)
            End With

            Dim fileHelper As New FileHelper
            Dim completed As Integer = fileHelper.writeToFile(fileName, sb.ToString)
            If completed = 1 Then
                System.Diagnostics.Process.Start("file:///" & m_Constant.FILE_FOLDER + fileName)
                'wbResult.Navigate(")
            End If

        End If
    End Sub

    Public Sub writeLabel(ByRef familyTable As SortedList, ByVal labelOption As Integer)
        Dim fileName As String = m_Constant.LABEL_SURVEY
        Dim sb As StringBuilder = New StringBuilder
        With sb
            .Append(printHeadHtml(m_Constant.PAPER_LETTER_PORTRAIT, True, labelOption))
            .Append("<body class='" + m_Constant.PAPER_LETTER_PORTRAIT + "'>" + vbCrLf)

            Dim keyList As IList = familyTable.GetKeyList()
            For i As Integer = 0 To keyList.Count - 1
                Dim j As Integer = i Mod m_Constant.AVERY_5160_MAX
                Dim k As Integer = i Mod 3

                If i = 0 Or j = 0 Then
                    .Append("<section class='sheet padding-8mm-2'>" + vbCrLf)
                    .Append("<table>" + vbCrLf)
                End If

                If i = 0 Or k = 0 Then
                    .Append("<tr>" + vbCrLf)
                End If

                Dim headName As String() = keyList(i).ToString.Split(CType("-", Char()))
                .Append("<td>" + headName(0) + "</td>" + vbCrLf)
                .Append("<td class='separator' />" + vbCrLf)

                If k = 2 Then
                    .Append("</tr>")
                End If

                If j = AVERY_5160_MAX - 1 Or i = keyList.Count - 1 Then
                    .Append("</table>" + vbCrLf)
                    .Append("</section>" + vbCrLf)
                End If
            Next

            .Append("</body>" + vbCrLf)
            .Append("</html>" + vbCrLf)
        End With

        Dim fileHelper As New FileHelper
        Dim completed As Integer = fileHelper.writeToFile(fileName, sb.ToString)
        If completed = 1 Then
            System.Diagnostics.Process.Start("file:///" & m_Constant.FILE_FOLDER + fileName)
        End If
        'fileHelper.convertHtmlToPDF(m_Constant.FILE_SURVEY, sb.ToString, m_Constant.ROOT_FOLDER + "\" + fileName)
        'System.Diagnostics.Process.Start(m_Constant.ROOT_FOLDER + "\" + fileName.Replace(m_Constant.FILE_EXTENSION_HTML, m_Constant.FILE_EXTENSION_PDF))

    End Sub

    Public Sub writeDirectoryBook(ByVal whichWindow As Integer)
        Dim ugResult As UltraWinGrid.UltraGrid = Nothing

        If whichWindow = 0 Then
            ugResult = mdiTKPC.ugResult
        End If

        Dim rows As Integer = ugResult.Rows.Count

        If rows > 0 Then
            Dim fileName As String = "tkpc_directory_book.html"
            Dim sb As StringBuilder = New StringBuilder
            Dim count As Integer = 0
            Dim familySeq As Integer = 0
            Dim familyDirEnt As TKPC.Entity.FamilyDirEnt = Nothing
            Dim familyList As SortedList = New SortedList
            Dim familySortKey As String = ""
            Dim familySpouseNames As String = ""
            Dim familyKidNames As String = ""
            Dim familyPhone1 As String = ""
            Dim familyPhone2 As String = ""
            Dim familyEmail As String = ""
            Dim familyEngName1 As String = ""
            Dim familyEngName2 As String = ""
            Dim familyAddress As String = ""
            Dim familyCity As String = ""

            For i As Integer = 0 To rows - 1
                Dim lvl As String = ugResult.Rows(i).Cells(0).Text

                If lvl = m_Constant.FAMILY_LEVEL_0 Or lvl = m_Constant.FAMILY_LEVEL_1 Then
                    familySpouseNames = ""
                    familyKidNames = ""
                    familyPhone1 = ""
                    familyPhone2 = ""
                    familyEmail = ""
                    familyEngName1 = ""
                    familyEngName2 = ""
                    familyAddress = ""
                    familyCity = ""
                    familyDirEnt = New TKPC.Entity.FamilyDirEnt
                    familySeq = 0
                End If

                Dim totalFamily As Integer = CType(ugResult.Rows(i).Cells(24).Text, Integer)
                Dim personId As String = ugResult.Rows(i).Cells(1).Text
                Dim koreanName As String = ugResult.Rows(i).Cells(2).Text
                Dim lastName As String = ugResult.Rows(i).Cells(3).Text
                Dim firstName As String = ugResult.Rows(i).Cells(4).Text
                Dim title As String = ugResult.Rows(i).Cells(8).Text
                Dim cell As String = ugResult.Rows(i).Cells(9).Text
                Dim home As String = ugResult.Rows(i).Cells(10).Text
                Dim email As String = ugResult.Rows(i).Cells(12).Text
                Dim type As String = ugResult.Rows(i).Cells(13).Text
                Dim address As String = ugResult.Rows(i).Cells(14).Text
                Dim city As String = ugResult.Rows(i).Cells(15).Text

                If type = m_Constant.FAMILY_TYPE_HEAD Then
                    familySpouseNames += koreanName + ","
                    familyPhone1 = home
                    familyPhone2 = cell
                    familyEmail = email
                    familyAddress = address
                    familyCity = city
                    familyEngName1 = lastName.ToUpper + ", " + firstName.ToUpper
                    familySortKey = koreanName + "-" + personId

                ElseIf type = m_Constant.FAMILY_TYPE_SPOUSE Then
                    familySpouseNames += koreanName + ","
                    familyEngName2 = lastName.ToUpper + ", " + firstName.ToUpper
                Else
                    If String.IsNullOrEmpty(koreanName) = False Then
                        familyKidNames += koreanName.Substring(1, koreanName.Length - 1) + ","
                    Else
                        familyKidNames += ","
                    End If
                End If

                If totalFamily = familySeq + 1 Then
                    If familySpouseNames.EndsWith(",") Then
                        familySpouseNames = familySpouseNames.Substring(0, familySpouseNames.Length - 1)
                    End If

                    If String.IsNullOrEmpty(familyKidNames) = False And familyKidNames.EndsWith(",") Then
                        familyKidNames = " (" + familyKidNames.Substring(0, familyKidNames.Length - 1) + ")"
                    End If

                    familyDirEnt.koreanNames = familySpouseNames + familyKidNames
                    familyDirEnt.engName1 = familyEngName1
                    familyDirEnt.engName2 = familyEngName2
                    familyDirEnt.address = familyAddress
                    familyDirEnt.city = familyCity
                    familyDirEnt.email = familyEmail

                    If familyList.ContainsKey(familySortKey) = False Then
                        familyList.Add(familySortKey, familyDirEnt)
                    End If
                End If

                familySeq += 1
            Next

            Dim dirRows As Integer = 4
            Dim dirRowsHeightCss As String = "height:650px"
            Dim modDivider As Integer = 8
            Dim halfPage As Integer = 3
            Dim endPage As Integer = 7

            If dirRows = 5 Then
                modDivider = 10
                halfPage = 4
                endPage = 9
            End If

            With sb
                .Append(m_Global.printHeadHtml(m_Constant.PAPER_LETTER_LANDSCAPE, False, -1))
                For Each familyObj As DictionaryEntry In familyList
                    Dim oneFamily As TKPC.Entity.FamilyDirEnt = CType(familyObj.Value, TKPC.Entity.FamilyDirEnt)
                    Dim wherePage As Integer = count Mod modDivider

                    If wherePage = 0 Then
                        .Append("<page size='letter' layout='landscape'>" + vbCrLf)
                        .Append("<body>" + vbCrLf)
                        .Append("<div id='textbox'>" + vbCrLf)
                        .Append("<p style='text-align:center; font-size:24px'>토론토 한인 장로 교회</p>" + vbCrLf)
                        .Append("</div>" + vbCrLf)
                        .Append("<table style='margin:0 auto; width: 95%;'>" + vbCrLf)
                        .Append("<tr>" + vbCrLf)
                        .Append("<td style='width:49%'>" + vbCrLf)
                        .Append("<table class='table table-bordered' style='margin-bottom:0px; " + dirRowsHeightCss + "'>" + vbCrLf)
                    End If

                    If dirRows = 5 Then
                        dirRowsHeightCss = ""
                    End If

                    .Append("<tr width='20%' height='155px'>" + vbCrLf)
                    .Append("<td style='text-align:center'><img src='./template/no_image.jpg' style='width:170px; height:110px'/></td>" + vbCrLf)
                    .Append("<td style='vertical-align:top;'>" + vbCrLf)
                    .Append("   <table style='margin-left:5px;'>" + vbCrLf)
                    .Append("   <tr>" + vbCrLf)
                    .Append("       <td><span style='font-weight:bold'>" + oneFamily.koreanNames + "</span></td>" + vbCrLf)
                    .Append("   </tr>" + vbCrLf)
                    .Append("   <tr>" + vbCrLf)
                    .Append("       <td>" + oneFamily.engName1.ToUpper + "</td>" + vbCrLf)
                    .Append("   </tr>" + vbCrLf)
                    .Append("   <tr>" + vbCrLf)
                    .Append("       <td>" + oneFamily.engName2.ToUpper + "</td>" + vbCrLf)
                    .Append("   </tr>" + vbCrLf)
                    .Append("   <tr>" + vbCrLf)
                    .Append("       <td>&nbsp;</td>" + vbCrLf)
                    .Append("   </tr>" + vbCrLf)
                    .Append("   <tr>" + vbCrLf)
                    .Append("       <td>" + oneFamily.address.ToUpper + "</td>" + vbCrLf)
                    .Append("   </tr>" + vbCrLf)
                    .Append("   <tr>" + vbCrLf)
                    .Append("       <td>" + oneFamily.city.ToUpper + "</td>" + vbCrLf)
                    .Append("   </tr>" + vbCrLf)
                    .Append("   <tr>" + vbCrLf)
                    .Append("       <td><span class='glyphicon glyphicon-phone' aria-hidden='true'></span>&nbsp;&nbsp;" + oneFamily.phone1 + "<span class='glyphicon glyphicon-phone-alt' aria-hidden='true'></span>&nbsp;&nbsp;" + oneFamily.phone2 + "</td>" + vbCrLf)
                    .Append("   </tr>" + vbCrLf)
                    .Append("   <tr>" + vbCrLf)
                    .Append("       <td><span class='glyphicon glyphicon-envelope' aria-hidden='true'></span>&nbsp;&nbsp;" + oneFamily.email + "</td>" + vbCrLf)
                    .Append("   </tr>" + vbCrLf)
                    .Append("   </table>" + vbCrLf)
                    .Append("</td>" + vbCrLf)
                    .Append("</tr>" + vbCrLf)



                    If wherePage = 3 Then
                        .Append("</table>" + vbCrLf)
                        .Append("</td>" + vbCrLf)
                        .Append("<td style='width:2%'></td>" + vbCrLf)
                        .Append("<td style='width:49%'>" + vbCrLf)
                        .Append("<table class='table table-bordered' style='margin-bottom:0px; " + dirRowsHeightCss + "'>" + vbCrLf)
                    End If

                    If wherePage = 7 Or count = familyList.Count - 1 Then
                        If wherePage <> 7 Then
                            Dim moreBox As Integer = 7 - wherePage
                            For k As Integer = 0 To moreBox - 1
                                .Append("<tr width='20%' height='155px'>" + vbCrLf)
                                .Append("<td style='text-align:center'><img src='./template/no_image.jpg' style='width:170px; height:110px'/></td>" + vbCrLf)
                                .Append("<td style='vertical-align:top;'>" + vbCrLf)
                                .Append("   <table style='margin-left:5px;'>" + vbCrLf)
                                .Append("   <tr>" + vbCrLf)
                                .Append("       <td><span style='font-weight:bold'></span></td>" + vbCrLf)
                                .Append("   </tr>" + vbCrLf)
                                .Append("   <tr>" + vbCrLf)
                                .Append("       <td></td>" + vbCrLf)
                                .Append("   </tr>" + vbCrLf)
                                .Append("   <tr>" + vbCrLf)
                                .Append("       <td></td>" + vbCrLf)
                                .Append("   </tr>" + vbCrLf)
                                .Append("   <tr>" + vbCrLf)
                                .Append("       <td>&nbsp;</td>" + vbCrLf)
                                .Append("   </tr>" + vbCrLf)
                                .Append("   <tr>" + vbCrLf)
                                .Append("       <td></td>" + vbCrLf)
                                .Append("   </tr>" + vbCrLf)
                                .Append("   <tr>" + vbCrLf)
                                .Append("       <td></td>" + vbCrLf)
                                .Append("   </tr>" + vbCrLf)
                                .Append("   <tr>" + vbCrLf)
                                .Append("       <td>&nbsp;&nbsp;&nbsp;&nbsp;</td>" + vbCrLf)
                                .Append("   </tr>" + vbCrLf)
                                .Append("   <tr>" + vbCrLf)
                                .Append("       <td>&nbsp;&nbsp;</td>" + vbCrLf)
                                .Append("   </tr>" + vbCrLf)
                                .Append("   </table>" + vbCrLf)
                                .Append("</td>" + vbCrLf)
                                .Append("</tr>" + vbCrLf)
                            Next
                        End If

                        .Append("</table>" + vbCrLf)
                        .Append("</td>" + vbCrLf)
                        .Append("</tr>" + vbCrLf)
                        .Append("<tr>" + vbCrLf)
                        .Append("<td style='text-align:center'>2-1</td>" + vbCrLf)
                        .Append("<td>&nbsp;</td>" + vbCrLf)
                        .Append("<td style='text-align:center'>2-1</td>" + vbCrLf)
                        .Append("</tr>" + vbCrLf)
                        .Append("</table>" + vbCrLf)
                        .Append("</body>" + vbCrLf)
                        .Append("</page>" + vbCrLf)
                    End If
                    count += 1
                Next
                .Append("</html>" + vbCrLf)
                'End If
            End With

            Dim fileHelper As New FileHelper
            Dim completed As Integer = fileHelper.writeToFile(fileName, sb.ToString)
            If completed = 1 Then
                System.Diagnostics.Process.Start("file:///" & m_Constant.FILE_FOLDER + fileName)
                'wbResult.Navigate(")
            End If

        End If
    End Sub
    Public Sub printNameTag()

        Dim sb As StringBuilder = New StringBuilder
        With sb
            .Append("<!DOCTYPE html> " + vbCrLf)
            .Append("<html> " + vbCrLf)
            .Append("<meta charset='utf-8'> " + vbCrLf)
            .Append("<head> " + vbCrLf)
            .Append("<link rel='stylesheet' href='https://cdnjs.cloudflare.com/ajax/libs/normalize/7.0.0/normalize.min.css'> " + vbCrLf)
            .Append("<link rel='stylesheet' href='https://cdnjs.cloudflare.com/ajax/libs/paper-css/0.4.1/paper.css'> " + vbCrLf)
            .Append("<link rel='stylesheet' href='https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css'> " + vbCrLf)
            .Append("<link rel='stylesheet' href='http://fonts.googleapis.com/earlyaccess/nanummyeongjo.css'> " + vbCrLf)
            .Append("<link rel='stylesheet' href='http://fonts.googleapis.com/earlyaccess/nanumpenscript.css'> " + vbCrLf)
            .Append("<link rel='stylesheet' href='http://fonts.googleapis.com/earlyaccess/jejuhallasan.css'> " + vbCrLf)
            .Append("<link rel='stylesheet' href='http://fonts.googleapis.com/earlyaccess/jejumyeongjo.css'> " + vbCrLf)
            .Append("<link rel='stylesheet' href='http://fonts.googleapis.com/earlyaccess/notosanskr.css'> " + vbCrLf)
            .Append("<link rel='stylesheet' href='http://fonts.googleapis.com/earlyaccess/nanumbrushscript.css'> " + vbCrLf)
            .Append("<link href='https://fonts.googleapis.com/css?family=Pangolin|Poppins|Candal|Hind+Guntur:700|Love+Ya+Like+A+Sister|Fontdiner+Swanky|Indie+Flower|Bad+Script' rel='stylesheet'> " + vbCrLf)
            .Append("<style> " + vbCrLf)
            .Append(".e_box { " + vbCrLf)
            .Append("		border: 2px dashed black; " + vbCrLf)
            .Append("		height: 280px; " + vbCrLf)
            .Append("		width: 440px; " + vbCrLf)
            .Append("} " + vbCrLf)
            .Append(".k_box { " + vbCrLf)
            .Append("		border: 2px dashed black; " + vbCrLf)
            .Append("		height: 280px; " + vbCrLf)
            .Append("		width: 440px; " + vbCrLf)
            .Append("} " + vbCrLf)
            .Append(".e_name { " + vbCrLf)
            .Append("		font-size: 75px; " + vbCrLf)
            .Append("		font-weight: bold; " + vbCrLf)
            .Append("		position: relative; " + vbCrLf)
            .Append("		top: 1%; " + vbCrLf)
            .Append("		text-align: center; " + vbCrLf)
            .Append("		font-family: 'Nanum Myeongjo', serif; " + vbCrLf)
            .Append("		margin-bottom: 0px; " + vbCrLf)
            .Append("} " + vbCrLf)
            .Append(".e_nameTitle { " + vbCrLf)
            .Append("		font-size: 40px; " + vbCrLf)
            .Append("		position: relative; " + vbCrLf)
            .Append("		/*font-family: 'Nanum Myeongjo', serif;*/ " + vbCrLf)
            .Append("		font-family: 'Indie Flower', cursive; " + vbCrLf)
            .Append("		/*font-family: 'Hind Guntur', sans-serif;*/ " + vbCrLf)
            .Append("} " + vbCrLf)
            .Append(".e_firstname { " + vbCrLf)
            .Append("		font-size: 68px; " + vbCrLf)
            .Append("		font-weight: bold; " + vbCrLf)
            .Append("		position: relative; " + vbCrLf)
            .Append("		top: 1%; " + vbCrLf)
            .Append("		text-align: center; " + vbCrLf)
            .Append("		/*font-family: 'Nanum Myeongjo', serif;*/ " + vbCrLf)
            .Append("		font-family: 'Candal', sans-serif; " + vbCrLf)
            .Append("		/*font-family: 'Hind Guntur', sans-serif;*/ " + vbCrLf)
            .Append("} " + vbCrLf)
            .Append(".e_lastname { " + vbCrLf)
            .Append("		font-size: 40px; " + vbCrLf)
            .Append("		font-weight: bold; " + vbCrLf)
            .Append("		text-align: center; " + vbCrLf)
            .Append("		/*font-family: 'Poppins', sans-serif;*/ " + vbCrLf)
            .Append("		/*font-family: 'Love Ya Like A Sister', cursive;*/ " + vbCrLf)
            .Append("		/*font-family: 'Fontdiner Swanky', cursive;*/ " + vbCrLf)
            .Append("		font-family: 'Indie Flower', cursive; " + vbCrLf)
            .Append("		margin: -10px 0 0 0; " + vbCrLf)
            .Append("} " + vbCrLf)
            .Append(".e_room { " + vbCrLf)
            .Append("		font-size: 37px; " + vbCrLf)
            .Append("		font-weight: bold; " + vbCrLf)
            .Append("		/*font-family: 'Nanum Pen Script', serif;*/ " + vbCrLf)
            .Append("		/*font-family: 'Noto Sans KR', serif; */ " + vbCrLf)
            .Append("		font-family: 'Poppins', sans-serif; " + vbCrLf)
            .Append("		text-align: left; " + vbCrLf)
            .Append("		margin-left: 5px; " + vbCrLf)
            .Append("} " + vbCrLf)
            .Append(".e_group1 { " + vbCrLf)
            .Append("		font-size: 25px; " + vbCrLf)
            .Append("		font-weight: bold; " + vbCrLf)
            .Append("		/*font-family: 'Nanum Pen Script', serif;*/ " + vbCrLf)
            .Append("		font-family: 'Jeju Myeongjo', serif; " + vbCrLf)
            .Append("		text-align: right; " + vbCrLf)
            .Append("		margin-right: 10px; " + vbCrLf)
            .Append("} " + vbCrLf)
            .Append(".e_group2 { " + vbCrLf)
            .Append("		font-size: 25px; " + vbCrLf)
            .Append("		font-weight: bold; " + vbCrLf)
            .Append("		/*font-family: 'Nanum Pen Script', serif;*/ " + vbCrLf)
            .Append("		/*font-family: 'Jeju Myeongjo', serif; */ " + vbCrLf)
            .Append("		font-family: 'Poppins', sans-serif; " + vbCrLf)
            .Append("		text-align: right; " + vbCrLf)
            .Append("		margin-right: 10px; " + vbCrLf)
            .Append("} " + vbCrLf)
            .Append(".e_title { " + vbCrLf)
            .Append("		font-size: 22px; " + vbCrLf)
            .Append("		font-weight: bold; " + vbCrLf)
            .Append("		/*font-family: 'Nanum Pen Script', serif;*/ " + vbCrLf)
            .Append("		font-family: 'Jeju Myeongjo', serif; " + vbCrLf)
            .Append("		text-align: center; " + vbCrLf)
            .Append("		margin-top: 10px; " + vbCrLf)
            .Append("} " + vbCrLf)
            .Append(".e_title2 { " + vbCrLf)
            .Append("		font-size: 25px; " + vbCrLf)
            .Append("		font-weight: bold; " + vbCrLf)
            .Append("		/*font-family: 'Nanum Myeongjo', serif;*/ " + vbCrLf)
            .Append("		/*font-family: 'Nanum Pen Script', serif;*/ " + vbCrLf)
            .Append("		/*font-family: 'Jeju Myeongjo', serif; */ " + vbCrLf)
            .Append("		font-family: 'Nanum Brush Script', serif; " + vbCrLf)
            .Append("		text-align: center; " + vbCrLf)
            .Append("		margin: -30px 0 0 -25px; " + vbCrLf)
            .Append("} " + vbCrLf)
            .Append(".e_title3 { " + vbCrLf)
            .Append("		font-size: 24px; " + vbCrLf)
            .Append("		font-weight: bold; " + vbCrLf)
            .Append("		/*font-family: 'Tahoma';*/ " + vbCrLf)
            .Append("		/*font-family: 'Pangolin', cursive;*/ " + vbCrLf)
            .Append("		font-family: 'Poppins', sans-serif; " + vbCrLf)
            .Append("		/*font-family: 'Nanum Myeongjo', serif;*/ " + vbCrLf)
            .Append("		/*font-family: 'Nanum Pen Script', serif;*/ " + vbCrLf)
            .Append("/*font-family: 'Jeju Myeongjo', serif;*/ " + vbCrLf)
            .Append("text-align: center; " + vbCrLf)
            .Append("margin: -45px 0 0 -25px; " + vbCrLf)
            .Append("} " + vbCrLf)
            .Append(".e_desc { " + vbCrLf)
            .Append("	font-size: 21px; " + vbCrLf)
            .Append("		font-weight: bold; " + vbCrLf)
            .Append("		font-family: 'Jeju Myeongjo', serif; " + vbCrLf)
            .Append("		text-align: left; " + vbCrLf)
            .Append("		padding: 5px 0px 3px 10px; " + vbCrLf)
            .Append("} " + vbCrLf)
            .Append(".e_desc2 { " + vbCrLf)
            .Append("		font-size: 21px; " + vbCrLf)
            .Append("		font-weight: bold; " + vbCrLf)
            .Append("		font-family: 'Jeju Myeongjo', serif; " + vbCrLf)
            .Append("		text-align: left; " + vbCrLf)
            .Append("		padding: 0px 5px 3px 33px; " + vbCrLf)
            .Append("} " + vbCrLf)
            .Append(".e_flipdiv { " + vbCrLf)
            .Append("		-moz-transform: rotate(180deg); " + vbCrLf)
            .Append("		-webkit-transform: rotate(180deg); " + vbCrLf)
            .Append("		-ms-transform: rotate(180deg); " + vbCrLf)
            .Append("		-o-transform: rotate(180deg); " + vbCrLf)
            .Append("		transform: rotate(180deg); " + vbCrLf)
            .Append("} " + vbCrLf)
            .Append(".e_tag { " + vbCrLf)
            .Append("		margin: auto; " + vbCrLf)
            .Append("} " + vbCrLf)
            .Append(".k_tag { " + vbCrLf)
            .Append("		margin: auto; " + vbCrLf)
            .Append("} " + vbCrLf)
            .Append(".e_logo { " + vbCrLf)
            .Append("		width: 80px; " + vbCrLf)
            .Append("		height: 80px; " + vbCrLf)
            .Append("} " + vbCrLf)
            .Append(".e_catchPhrase { " + vbCrLf)
            .Append("		position: relative; " + vbCrLf)
            .Append("		margin-left: 25px; " + vbCrLf)
            .Append("		font-size: 22px; " + vbCrLf)
            .Append("		font-weight: bold; " + vbCrLf)
            .Append("		/*font-family: 'Nanum Pen Script', serif;*/ " + vbCrLf)
            .Append("		font-family: 'Jeju Myeongjo', serif; " + vbCrLf)
            .Append("		text-align: center; " + vbCrLf)
            .Append("} " + vbCrLf)
            .Append(".e_game { " + vbCrLf)
            .Append("		position: absolute; " + vbCrLf)
            .Append("		bottom: 0; " + vbCrLf)
            .Append("		right: 0; " + vbCrLf)
            .Append("		font-weight: bold; " + vbCrLf)
            .Append("       margin-right: 20px; " + vbCrLf)
            .Append("} " + vbCrLf)
            .Append(".k_name { " + vbCrLf)
            .Append("		font-size: 83px; " + vbCrLf)
            .Append("		font-weight: bold; " + vbCrLf)
            .Append("		position: relative; " + vbCrLf)
            .Append("		top: 2%; " + vbCrLf)
            .Append("		text-align: center; " + vbCrLf)
            .Append("		font-family: 'Nanum Myeongjo', serif; " + vbCrLf)
            .Append("		margin-bottom: 15px; " + vbCrLf)
            .Append("} " + vbCrLf)
            .Append(".k_room { " + vbCrLf)
            .Append("		font-size: 37px; " + vbCrLf)
            .Append("		font-weight: bold; " + vbCrLf)
            .Append("		/*font-family: 'Nanum Pen Script', serif;*/ " + vbCrLf)
            .Append("		/*font-family: 'Noto Sans KR', serif; */ " + vbCrLf)
            .Append("		font-family: 'Poppins', sans-serif; " + vbCrLf)
            .Append("		text-align: left; " + vbCrLf)
            .Append("		margin-left: 5px; " + vbCrLf)
            .Append("} " + vbCrLf)
            .Append(".k_group1 { " + vbCrLf)
            .Append("		font-size: 22px; " + vbCrLf)
            .Append("		font-weight: bold; " + vbCrLf)
            .Append("		/*font-family: 'Nanum Pen Script', serif;*/ " + vbCrLf)
            .Append("		font-family: 'Jeju Myeongjo', serif; " + vbCrLf)
            .Append("		text-align: right; " + vbCrLf)
            .Append("		margin-right: 10px; " + vbCrLf)
            .Append("} " + vbCrLf)
            .Append(".k_group2 { " + vbCrLf)
            .Append("		font-size: 22px; " + vbCrLf)
            .Append("		font-weight: bold; " + vbCrLf)
            .Append("		/*font-family: 'Nanum Pen Script', serif;*/ " + vbCrLf)
            .Append("		font-family: 'Jeju Myeongjo', serif; " + vbCrLf)
            .Append("		text-align: right; " + vbCrLf)
            .Append("		margin-right: 10px; " + vbCrLf)
            .Append("} " + vbCrLf)
            .Append(".k_title { " + vbCrLf)
            .Append("		font-size: 25px; " + vbCrLf)
            .Append("		font-weight: bold; " + vbCrLf)
            .Append("		/*font-family: 'Nanum Pen Script', serif;*/ " + vbCrLf)
            .Append("		font-family: 'Jeju Myeongjo', serif; " + vbCrLf)
            .Append("		text-align: center; " + vbCrLf)
            .Append("		margin-top: 10px; " + vbCrLf)
            .Append("} " + vbCrLf)
            .Append("u { " + vbCrLf)
            .Append("		text-decoration: none; " + vbCrLf)
            .Append("		border-bottom: 2px solid black; " + vbCrLf)
            .Append("} " + vbCrLf)
            .Append(".k_title2 { " + vbCrLf)
            .Append("		font-size: 25px; " + vbCrLf)
            .Append("		font-weight: bold; " + vbCrLf)
            .Append("		/*font-family: 'Nanum Myeongjo', serif;*/ " + vbCrLf)
            .Append("		/*font-family: 'Nanum Pen Script', serif;*/ " + vbCrLf)
            .Append("		/*font-family: 'Jeju Myeongjo', serif; */ " + vbCrLf)
            .Append("		font-family: 'Nanum Brush Script', serif; " + vbCrLf)
            .Append("		text-align: center; " + vbCrLf)
            .Append("		margin: -30px 0 0 -25px; " + vbCrLf)
            .Append("} " + vbCrLf)
            .Append(".k_title3 { " + vbCrLf)
            .Append("		font-size: 24px; " + vbCrLf)
            .Append("		font-weight: bold; " + vbCrLf)
            .Append("		/*font-family: 'Tahoma';*/ " + vbCrLf)
            .Append("		/*font-family: 'Pangolin', cursive;*/ " + vbCrLf)
            .Append("		font-family: 'Poppins', sans-serif; " + vbCrLf)
            .Append("		/*font-family: 'Nanum Myeongjo', serif;*/ " + vbCrLf)
            .Append("		/*font-family: 'Nanum Pen Script', serif;*/ " + vbCrLf)
            .Append("		/*font-family: 'Jeju Myeongjo', serif;*/ " + vbCrLf)
            .Append("		/*font-family: 'Nanum Brush Script', serif;*/ " + vbCrLf)
            .Append("		text-align: center; " + vbCrLf)
            .Append("		margin: -45px 0 0 -25px; " + vbCrLf)
            .Append("} " + vbCrLf)
            .Append(".k_desc { " + vbCrLf)
            .Append("		font-size: 24px; " + vbCrLf)
            .Append("		font-weight: bold; " + vbCrLf)
            .Append("		font-family: 'Jeju Myeongjo', serif; " + vbCrLf)
            .Append("		text-align: left; " + vbCrLf)
            .Append("		padding: 0px 0px 3px 10px; " + vbCrLf)
            .Append("} " + vbCrLf)
            .Append(".k_desc2 { " + vbCrLf)
            .Append("		font-size: 24px; " + vbCrLf)
            .Append("		font-weight: bold; " + vbCrLf)
            .Append("		font-family: 'Jeju Myeongjo', serif; " + vbCrLf)
            .Append("		text-align: left; " + vbCrLf)
            .Append("		padding: 0px 0px 3px 28px; " + vbCrLf)
            .Append("} " + vbCrLf)
            .Append(".k_flipdiv { " + vbCrLf)
            .Append("		-moz-transform: rotate(180deg); " + vbCrLf)
            .Append("		-webkit-transform: rotate(180deg); " + vbCrLf)
            .Append("		-ms-transform: rotate(180deg); " + vbCrLf)
            .Append("		-o-transform: rotate(180deg); " + vbCrLf)
            .Append("		transform: rotate(180deg); " + vbCrLf)
            .Append("} " + vbCrLf)
            .Append(".k_logo { " + vbCrLf)
            .Append("		width: 80px; " + vbCrLf)
            .Append("		height: 80px; " + vbCrLf)
            .Append("} " + vbCrLf)
            .Append(".k_catchPhrase { " + vbCrLf)
            .Append("		position: relative; " + vbCrLf)
            .Append("		margin-left: 25px; " + vbCrLf)
            .Append("		font-size: 22px; " + vbCrLf)
            .Append("		font-weight: bold; " + vbCrLf)
            .Append("		/*font-family: 'Nanum Pen Script', serif;*/ " + vbCrLf)
            .Append("		font-family: 'Jeju Myeongjo', serif; " + vbCrLf)
            .Append("		text-align: center; " + vbCrLf)
            .Append("} " + vbCrLf)
            .Append(".k_game { " + vbCrLf)
            .Append("		position: absolute; " + vbCrLf)
            .Append("		bottom: 0; " + vbCrLf)
            .Append("		right: 0; " + vbCrLf)
            .Append("		font-weight: bold; " + vbCrLf)
            .Append("        margin-right: 20px; " + vbCrLf)
            .Append("} " + vbCrLf)
            .Append("</style> " + vbCrLf)
            .Append("</head> " + vbCrLf)
            .Append("<body class='A3 landscape'> " + vbCrLf)
            .Append("<section class='sheet padding-5mm'> " + vbCrLf)
            .Append("<table> " + vbCrLf)
            .Append("<tr> " + vbCrLf)
            .Append("		<td> " + vbCrLf)
            .Append("			<div class='k_box'> " + vbCrLf)
            .Append("			<div class='row' style='margin-left: 5px;'> " + vbCrLf)
            .Append("				<table width='96%'> " + vbCrLf)
            .Append("				<tr> " + vbCrLf)
            .Append("					<td rowspan='2' width='81px'><div class='k_logo'><img src='logo.jpeg' class='k_logo'></div></td> " + vbCrLf)
            .Append("					<td><div class='k_title2'>TKPC 50주년 전교인 비젼 수련회</div></td> " + vbCrLf)
            .Append("				</tr> " + vbCrLf)
            .Append("				<tr> " + vbCrLf)
            .Append("					<td><div class='k_title3'>One Body One Family</div></td> " + vbCrLf)
            .Append("				</tr> " + vbCrLf)
            .Append("				</table> " + vbCrLf)
            .Append("			</div> " + vbCrLf)
            .Append("		<div class='k_name'>배덕출</div> " + vbCrLf)
            .Append("			<div class='row' style='margin-left: 10px;'> " + vbCrLf)
            .Append("				<table style='width:96%; margin-bottom:0px'> " + vbCrLf)
            .Append("				<tr> " + vbCrLf)
            .Append("					<td rowspan='2' width='50%'><div class='k_room'>Vallee501</div></td> " + vbCrLf)
            .Append("					<td width='50%'><div class='k_group1'>소그룹: 60-1M</div></td> " + vbCrLf)
            .Append("				</tr> " + vbCrLf)
            .Append("				<tr> " + vbCrLf)
            .Append("					<td><div class='k_group2'>세족식: 67</div></td> " + vbCrLf)
            .Append("				</tr> " + vbCrLf)
            .Append("				</table> " + vbCrLf)
            .Append("			</div> " + vbCrLf)
            .Append("			</div> " + vbCrLf)
            .Append("		</td> " + vbCrLf)
            .Append("		<td> " + vbCrLf)
            .Append("			<div class='k_box'> " + vbCrLf)
            .Append("			<div class='row' style='margin-left: 5px;'> " + vbCrLf)
            .Append("				<table width='96%'> " + vbCrLf)
            .Append("				<tr> " + vbCrLf)
            .Append("					<td rowspan='2' width='81px'><div class='k_logo'><img src='logo.jpeg' class='k_logo'></div></td> " + vbCrLf)
            .Append("					<td><div class='k_title2'>TKPC 50주년 전교인 비젼 수련회</div></td> " + vbCrLf)
            .Append("				</tr> " + vbCrLf)
            .Append("				<tr> " + vbCrLf)
            .Append("					<td><div class='k_title3'>One Body One Family</div></td> " + vbCrLf)
            .Append("				</tr> " + vbCrLf)
            .Append("				</table> " + vbCrLf)
            .Append("			</div> " + vbCrLf)
            .Append("			<div class='k_name'>배원철</div> " + vbCrLf)
            .Append("			<div class='row' style='margin-left: 10px;'> " + vbCrLf)
            .Append("				<table style='width:96%; margin-bottom:0px'> " + vbCrLf)
            .Append("				<tr> " + vbCrLf)
            .Append("					<td rowspan='2' width='50%'><div class='k_room'>Vallee502</div></td> " + vbCrLf)
            .Append("					<td width='50%'><div class='k_group1'>소그룹: 50-3M</div></td> " + vbCrLf)
            .Append("				</tr> " + vbCrLf)
            .Append("				<tr> " + vbCrLf)
            .Append("					<td><div class='k_group2'>세족식: 2</div></td> " + vbCrLf)
            .Append("				</tr> " + vbCrLf)
            .Append("				</table> " + vbCrLf)
            .Append("			</div> " + vbCrLf)
            .Append("			</div> " + vbCrLf)
            .Append("		</td> " + vbCrLf)
            .Append("</tr> " + vbCrLf)
            .Append("<tr> " + vbCrLf)
            .Append("		<td> " + vbCrLf)
            .Append("			<div class='k_box k_flipdiv'> " + vbCrLf)
            .Append("				<div class='k_game'>노랑</div> " + vbCrLf)
            .Append("				<div class='k_desc' style='margin-top:10px'>* 귀가전 방 열쇠와 명찰은 반납해주세요.<br>(방 열쇠 분실시 본인부담 230불입니다.)</div> " + vbCrLf)
            .Append("				<div class='k_desc'>* 외출시 Floor Leader에게 꼭 알려주세요.</div> " + vbCrLf)
            .Append("				<div class='k_desc'>&nbsp;&nbsp;&nbsp;<u>위급상황시 연락처</u></div> " + vbCrLf)
            .Append("				<div class='k_desc2'>강영태: 416-450-1012</div> " + vbCrLf)
            .Append("				<div class='k_desc2'>장성훈: 647-457-0191</div> " + vbCrLf)
            .Append("				<div class='k_desc2' style='font-size:20px;'>Floor Leader: 김재호 (511호), 김동규 (518호)</div> " + vbCrLf)
            .Append("			</div> " + vbCrLf)
            .Append("		</td> " + vbCrLf)
            .Append("		<td> " + vbCrLf)
            .Append("			<div class='k_box k_flipdiv'> " + vbCrLf)
            .Append("				<div class='k_game'>파랑</div> " + vbCrLf)
            .Append("				<div class='k_desc' style='margin-top:10px'>* 귀가전 방 열쇠와 명찰은 반납해주세요.<br>(방 열쇠 분실시 본인부담 230불입니다.)</div> " + vbCrLf)
            .Append("				<div class='k_desc'>* 외출시 Floor Leader에게 꼭 알려주세요.</div> " + vbCrLf)
            .Append("				<div class='k_desc'>&nbsp;&nbsp;&nbsp;<u>위급상황시 연락처</u></div> " + vbCrLf)
            .Append("				<div class='k_desc2'>강영태: 416-450-1012</div> " + vbCrLf)
            .Append("				<div class='k_desc2'>장성훈: 647-457-0191</div> " + vbCrLf)
            .Append("				<div class='k_desc2' style='font-size:20px;'>Floor Leader: 김재호 (511호), 김동규 (518호)</div> " + vbCrLf)
            .Append("			</div> " + vbCrLf)
            .Append("		</td> " + vbCrLf)
            .Append("</tr> " + vbCrLf)
            .Append("</table> " + vbCrLf)
            .Append("<table> " + vbCrLf)
            .Append("<tr> " + vbCrLf)
            .Append("		<td> " + vbCrLf)
            .Append("			<div class='k_box'> " + vbCrLf)
            .Append("			<div class='row' style='margin-left: 5px;'> " + vbCrLf)
            .Append("				<table width='96%'> " + vbCrLf)
            .Append("				<tr> " + vbCrLf)
            .Append("					<td rowspan='2' width='81px'><div class='k_logo'><img src='logo.jpeg' class='k_logo'></div></td> " + vbCrLf)
            .Append("					<td><div class='k_title2'>TKPC 50주년 전교인 비젼 수련회</div></td> " + vbCrLf)
            .Append("				</tr> " + vbCrLf)
            .Append("				<tr> " + vbCrLf)
            .Append("					<td><div class='k_title3'>One Body One Family</div></td> " + vbCrLf)
            .Append("				</tr> " + vbCrLf)
            .Append("				</table> " + vbCrLf)
            .Append("			</div> " + vbCrLf)
            .Append("			<div class='k_name'>배덕출</div> " + vbCrLf)
            .Append("			<div class='row' style='margin-left: 10px;'> " + vbCrLf)
            .Append("				<table style='width:96%; margin-bottom:0px'> " + vbCrLf)
            .Append("				<tr> " + vbCrLf)
            .Append("					<td rowspan='2' width='50%'><div class='k_room'>Vallee501</div></td> " + vbCrLf)
            .Append("					<td width='50%'><div class='k_group1'>소그룹: 60-1M</div></td> " + vbCrLf)
            .Append("				</tr> " + vbCrLf)
            .Append("				<tr> " + vbCrLf)
            .Append("					<td><div class='k_group2'>세족식: 67</div></td> " + vbCrLf)
            .Append("				</tr> " + vbCrLf)
            .Append("				</table> " + vbCrLf)
            .Append("			</div> " + vbCrLf)
            .Append("			</div> " + vbCrLf)
            .Append("		</td> " + vbCrLf)
            .Append("		<td> " + vbCrLf)
            .Append("			<div class='k_box'> " + vbCrLf)
            .Append("			<div class='row' style='margin-left: 5px;'> " + vbCrLf)
            .Append("				<table width='96%'> " + vbCrLf)
            .Append("				<tr> " + vbCrLf)
            .Append("					<td rowspan='2' width='81px'><div class='k_logo'><img src='logo.jpeg' class='k_logo'></div></td> " + vbCrLf)
            .Append("					<td><div class='k_title2'>TKPC 50주년 전교인 비젼 수련회</div></td> " + vbCrLf)
            .Append("				</tr> " + vbCrLf)
            .Append("				<tr> " + vbCrLf)
            .Append("					<td><div class='k_title3'>One Body One Family</div></td> " + vbCrLf)
            .Append("				</tr> " + vbCrLf)
            .Append("				</table> " + vbCrLf)
            .Append("			</div> " + vbCrLf)
            .Append("			<div class='k_name'>배원철</div> " + vbCrLf)
            .Append("			<div class='row' style='margin-left: 10px;'> " + vbCrLf)
            .Append("				<table style='width:96%; margin-bottom:0px'> " + vbCrLf)
            .Append("				<tr> " + vbCrLf)
            .Append("					<td rowspan='2' width='50%'><div class='k_room'>Vallee502</div></td> " + vbCrLf)
            .Append("					<td width='50%'><div class='k_group1'>소그룹: 50-3M</div></td> " + vbCrLf)
            .Append("				</tr> " + vbCrLf)
            .Append("				<tr> " + vbCrLf)
            .Append("					<td><div class='k_group2'>세족식: 2</div></td> " + vbCrLf)
            .Append("				</tr> " + vbCrLf)
            .Append("				</table> " + vbCrLf)
            .Append("			</div> " + vbCrLf)
            .Append("			</div> " + vbCrLf)
            .Append("		</td> " + vbCrLf)
            .Append("</tr> " + vbCrLf)
            .Append("<tr> " + vbCrLf)
            .Append("		<td> " + vbCrLf)
            .Append("			<div class='k_box k_flipdiv'> " + vbCrLf)
            .Append("				<div class='k_game'>노랑</div> " + vbCrLf)
            .Append("				<div class='k_desc' style='margin-top:10px'>* 귀가전 방 열쇠와 명찰은 반납해주세요.<br>(방 열쇠 분실시 본인부담 230불입니다.)</div> " + vbCrLf)
            .Append("				<div class='k_desc'>* 외출시 Floor Leader에게 꼭 알려주세요.</div> " + vbCrLf)
            .Append("				<div class='k_desc'>&nbsp;&nbsp;&nbsp;<u>위급상황시 연락처</u></div> " + vbCrLf)
            .Append("				<div class='k_desc2'>강영태: 416-450-1012</div> " + vbCrLf)
            .Append("				<div class='k_desc2'>장성훈: 647-457-0191</div> " + vbCrLf)
            .Append("				<div class='k_desc2' style='font-size:20px;'>Floor Leader: 김재호 (511호), 김동규 (518호)</div> " + vbCrLf)
            .Append("			</div> " + vbCrLf)
            .Append("		</td> " + vbCrLf)
            .Append("		<td> " + vbCrLf)
            .Append("			<div class='k_box k_flipdiv'> " + vbCrLf)
            .Append("				<div class='k_game'>파랑</div> " + vbCrLf)
            .Append("				<div class='k_desc' style='margin-top:10px'>* 귀가전 방 열쇠와 명찰은 반납해주세요.<br>(방 열쇠 분실시 본인부담 230불입니다.)</div> " + vbCrLf)
            .Append("				<div class='k_desc'>* 외출시 Floor Leader에게 꼭 알려주세요.</div> " + vbCrLf)
            .Append("				<div class='k_desc'>&nbsp;&nbsp;&nbsp;<u>위급상황시 연락처</u></div> " + vbCrLf)
            .Append("				<div class='k_desc2'>강영태: 416-450-1012</div> " + vbCrLf)
            .Append("				<div class='k_desc2'>장성훈: 647-457-0191</div> " + vbCrLf)
            .Append("				<div class='k_desc2' style='font-size:20px;'>Floor Leader: 김재호 (511호), 김동규 (518호)</div> " + vbCrLf)
            .Append("			</div> " + vbCrLf)
            .Append("		</td> " + vbCrLf)
            .Append("</tr> " + vbCrLf)
            .Append("</table> " + vbCrLf)
            .Append("</section> " + vbCrLf)
            .Append("</body> " + vbCrLf)
            .Append("</html> " + vbCrLf)
        End With
    End Sub

    Public Function formatDate(ByVal flag As Integer, ByVal dateStr As String, ByVal ddStr As String) As String
        formatDate = ""
        If String.IsNullOrEmpty(dateStr) = True Then
            Exit Function
        End If

        Dim splitDate As String() = dateStr.Split(CType("-", Char))
        If flag = 0 Then 'Korean way
            formatDate = splitDate(0) + "년" + CType(Integer.Parse(splitDate(1)), String) + "월" + CType(Integer.Parse(splitDate(2)), String) + "일"
        ElseIf flag = 1 Then '월단위
            Dim month As Integer = CType(splitDate(1), Integer)
            If ddStr <> "01" Then
                If DateTime.IsLeapYear(CType(splitDate(0), Integer)) Then
                    If month = 1 Or month = 3 Or month = 5 Or month = 7 Or month = 8 Or month = 10 Or month = 12 Then
                        ddStr = "31"
                    ElseIf month = 2 Then
                        ddStr = "29"
                    Else
                        ddStr = "30"
                    End If
                Else
                    If month = 1 Or month = 3 Or month = 5 Or month = 7 Or month = 8 Or month = 10 Or month = 12 Then
                        ddStr = "31"
                    ElseIf month = 2 Then
                        ddStr = "28"
                    Else
                        ddStr = "30"
                    End If
                End If
            End If

            formatDate = splitDate(0) + "-" + splitDate(1) + "-" + ddStr
        End If
        Return formatDate
    End Function


    Public Function compare(ByRef first As Integer, ByRef second As Integer) As Boolean
        If first <= second Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function compare2Dates(ByRef first As Date, ByRef second As Date) As Boolean
        If first <= second Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub setNavigatorBinding(ByRef ds As DataSet, ByVal flag As Integer, ByVal cardButtonVisibleFlag As Boolean, ByVal title As String)
        If ds IsNot Nothing Then
            Dim bs As BindingSource = New BindingSource

            bs.DataSource = ds.Tables(0)
            Dim totalNumberOfRows As Integer = ds.Tables(0).Rows.Count

            If flag = 0 Then
                If mdiTKPC.Visible = False Then
                    mdiTKPC.Visible = True
                End If

                mdiTKPC.bnFind.BindingSource = bs
                With mdiTKPC.ugFind
                    .DataSource = Nothing
                    .DataSource = bs
                    .Focus()
                End With

                If mdiTKPC.ugFind.Rows.Count > 0 Then
                    mdiTKPC.tsbtFindAll.Enabled = True
                    mdiTKPC.tsbtFindUnAll.Enabled = True
                    mdiTKPC.tsbtFindExcel.Enabled = True
                    mdiTKPC.tsbtFindText.Enabled = True
                    mdiTKPC.tsbtLabel.Visible = False
                    mdiTKPC.tsbtFindPrint.Visible = True
                Else
                    mdiTKPC.tsbtFindAll.Enabled = False
                    mdiTKPC.tsbtFindUnAll.Enabled = False
                    mdiTKPC.tsbtFindExcel.Enabled = False
                    mdiTKPC.tsbtFindText.Enabled = False
                    mdiTKPC.tsbtLabel.Visible = False
                    mdiTKPC.tsbtFindPrint.Visible = False
                End If
                'Dim r As DataRow
                'mdiTKPC.txtFind.AutoCompleteCustomSource.Clear()
                'For Each r In ds.Tables(0).Rows
                '    mdiTKPC.txtFind.AutoCompleteCustomSource.Add(CType(r.Item(0), String))
                '    mdiTKPC.txtFind.AutoCompleteCustomSource.Add(CType(r.Item(1), String))
                'Next

            ElseIf flag = 1 Then
                mdiTKPC.bnList.BindingSource = bs
                With mdiTKPC.ugResult
                    .Text = title
                    .DataSource = Nothing
                End With

                If totalNumberOfRows = 0 Then
                    mdiTKPC.tsbtExcel.Visible = False
                    mdiTKPC.tsbtCard.Visible = False
                    mdiTKPC.tsbtSurvey.Visible = False
                    mdiTKPC.tsbtLabel.Visible = False
                    mdiTKPC.tsbtPrint.Visible = False
                    MsgBox("데이터를 찾을 수 없습니다. 옮겨담기한 교우들을 삭제하거나 기간 옵션을 변경한 후 다시 시도해주십시오.", Nothing, "결과 없음")
                    Exit Sub
                Else
                    With mdiTKPC.ugResult
                        .DataSource = Nothing
                        .DataSource = bs
                        .Focus()
                        .Text = title
                    End With
                    mdiTKPC.tsbtExcel.Visible = True
                    mdiTKPC.tsbtLabel.Visible = False
                    mdiTKPC.tsbtPrint.Visible = True
                    'mdiTKPC.tsbtCard.Visible = cardButtonVisibleFlag
                    'mdiTKPC.tsbtSurvey.Visible = cardButtonVisibleFlag
                End If
            Else
                frmVisitInfoOption.bnList.BindingSource = bs
                With frmVisitInfoOption.ugResult
                    .DataSource = Nothing
                    .DataSource = bs
                    .DisplayLayout.Bands(0).Columns("spouses").Hidden = True
                    .DisplayLayout.Bands(0).Columns("person_id").Hidden = True
                    .Focus()
                End With

                If totalNumberOfRows = 0 Then
                    With frmVisitInfoOption
                        .tsbtCheckAll.Visible = False
                        .tsbtUncheckAll.Visible = False
                    End With
                End If
            End If
        Else
            If flag = 0 Then
                mdiTKPC.ugFind.Visible = False
            Else
                mdiTKPC.tsbtExcel.Visible = False
                mdiTKPC.tsbtLabel.Visible = False
                mdiTKPC.tsbtPrint.Visible = False
            End If
        End If
    End Sub
    Public Function compareString(ByRef str1 As String, ByRef str2 As String) As Integer
        Return String.Compare(str1, str2)
    End Function

    'Get the first day of the month
    Public Function FirstDayOfMonth(ByVal sourceDate As DateTime) As DateTime
        Return New DateTime(sourceDate.Year, sourceDate.Month, 1)
    End Function

    'Get the last day of the month
    Public Function LastDayOfMonth(ByVal sourceDate As DateTime) As DateTime
        Dim lastDay As DateTime = New DateTime(sourceDate.Year, sourceDate.Month, 1)
        Return lastDay.AddMonths(1).AddDays(-1)
    End Function

    Public Function getOfferingDates(ByVal flag As Integer, ByVal startDate As String, ByVal endDate As String) As List(Of String)
        getOfferingDates = New List(Of String)
        Dim offeringDAL As TKPC.DAL.OfferingDAL = TKPC.DAL.OfferingDAL.getInstance()
        Dim ds As DataSet = Nothing

        Try
            ds = offeringDAL.getOfferingDates(flag, startDate, endDate)

            If ds IsNot Nothing Then
                Dim dr As DataRow
                For Each dr In ds.Tables(0).Rows
                    getOfferingDates.Add(CType(dr(0), String))
                Next
            End If

        Catch ex As Exception
            Logger.LogError(ex.ToString)
        Finally
            offeringDAL = Nothing
            ds = Nothing
        End Try

        Return getOfferingDates
    End Function

    Public Function getOfferingYears(ByVal startDate As String, ByVal endDate As String) As String
        getOfferingYears = ""
        Dim startYear As Integer = CType(startDate, Date).Year
        Dim endYear As Integer = CType(endDate, Date).Year
        Dim tempYear As Integer = 0

        For i As Integer = startYear To endYear
            tempYear = startYear
            getOfferingYears += CType(tempYear, String)
            If tempYear <> endYear Then
                getOfferingYears += ","
            Else
                Exit For
            End If
            tempYear += 1
        Next
        Return getOfferingYears
    End Function

    Public Declare Function LoadKeyboardLayout Lib "user32" Alias "LoadKeyboardLayoutA" (
        ByVal pwszKLID As String, ByVal Flags As Long) As Long

    Public Declare Function UnloadKeyboardLayout Lib "user32" (
        ByVal HKL As Long) As Long

    Public Declare Function ActivateKeyboardLayout Lib "user32" (
        ByVal HKL As Long, ByVal Flags As Long) As Long

    Public Declare Function GetKeyboardLayout Lib "user32" (
        ByVal dwLayout As Long) As Long

    Public Declare Function GetKeyboardLayoutName Lib "user32" Alias "GetKeyboardLayoutNameA" (
        ByVal pwszKLID As String) As Long

    Public Function SwitchIME(ByRef hLayoutId As Long, ByRef KbdId As String) As Integer
        '//Loading Keyboard requires string identifier
        SwitchIME = 0
        hLayoutId = LoadKeyboardLayout(KbdId, 0)
        ActivateKeyboardLayout(hLayoutId, 0)
        If KbdId = KbdKr Then
            SwitchIME = 1
        End If
    End Function
End Module

Imports System.Collections.Generic
Imports System.Configuration
Imports System.Drawing
Imports System.Text

Public Class frmReceiptInfo
    Public flag As Integer = 0
    Public personId As String = ""
    Public offeringYear As String = ""
    Public lastName As String = ""
    Public firstName As String = ""
    Public street As String = ""
    Public city As String = ""
    Public postalcode As String = ""
    Public offeringNumber As String = ""

    Private Sub frmReceiptInfo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '<add key = "church_name" value="Toronto Korean Presbyterian Church"/>
        '<add key = "church_addr1" value="67 Scarsdale Rd."/>
        '<add key = "church_addr2" value="North York, ON. M3B 2R2"/>
        '<add key = "church_phone" value="416-477-5963"/>
        '<add key = "church_reg_no" value="82777 7970 RR0001"/>
        '<add key = "treasurer" value="Kim Sung Kyum"/>

        txtName.Text = ConfigurationManager.AppSettings(m_Constant.RECEIPT_CHURCH_NAME).ToString()
        txtAddr1.Text = ConfigurationManager.AppSettings(m_Constant.RECEIPT_CHURCH_ADDR1).ToString()
        txtAddr2.Text = ConfigurationManager.AppSettings(m_Constant.RECEIPT_CHURCH_ADDR2).ToString()
        txtPhone.Text = ConfigurationManager.AppSettings(m_Constant.RECEIPT_CHURCH_PHONE).ToString()
        txtRegNo.Text = ConfigurationManager.AppSettings(m_Constant.RECEIPT_CHURCH_REG_NO).ToString()
        txtTreasurer.Text = ConfigurationManager.AppSettings(m_Constant.RECEIPT_TREASURER).ToString()
        picSignature.ImageLocation = m_Constant.ROOT_FOLDER + "\images\receipt_signature.png"

        lblPath.Text = m_Constant.ROOT_FOLDER + "\images\receipt_signature.png"
    End Sub

    Private Sub writeOfferingReceipt()
        Dim sb As StringBuilder = New StringBuilder
        Me.Cursor = Cursors.WaitCursor
        Dim offeringDAL As TKPC.DAL.OfferingDAL = TKPC.DAL.OfferingDAL.getInstance()

        Dim cp As List(Of String) = Nothing

        If flag = 0 Then
            cp = mdiTKPC.readPersonIdFromDgChosen(0)
        Else
            cp = New List(Of String)
            cp.Add(personId)
        End If

        Dim personIds As String = ""
        If cp.Count > 0 Then
            personIds = String.Join(",", cp)
        End If

        Dim startYear As String = ""
        Dim endYear As String = ""

        If flag = 0 Then
            startYear = mdiTKPC.dtpCurrentStartDate.Text
            endYear = mdiTKPC.dtpCurrentEndDate.Text
        Else
            startYear = offeringYear
            endYear = offeringYear
        End If

        Dim ds As DataSet = offeringDAL.getOfferingReceipt(1, startYear, endYear, personIds, street)

        Try
            Dim rowCount As Integer = 0
            If flag = 0 Then
                mdiTKPC.ugResult.DataSource = Nothing
                mdiTKPC.ugResult.DataSource = ds.Tables(0)
                mdiTKPC.ugResult.Text = "헌금 영수증 리스트"
                rowCount = mdiTKPC.ugResult.Rows.Count
            Else
                rowCount = 1
            End If

            Dim selectedYear = startYear
            Dim fileName As String = "tkpc_" + selectedYear + "_donation.html"

            sb.Append("<!DOCTYPE html>" + vbCrLf)
            sb.Append("<html lang='en'>" + vbCrLf)
            sb.Append("<head>" + vbCrLf)
            sb.Append("<meta charset='utf-8'>" + vbCrLf)
            sb.Append("<title>" + selectedYear + "년도 헌금 영수증</title>" + vbCrLf)
            sb.Append("<link rel='stylesheet' href='../css/normalize.min.css'>" + vbCrLf)
            sb.Append("<link rel='stylesheet' href='../css/paper.css'>" + vbCrLf)
            sb.Append("<style>" + vbCrLf)
            sb.Append("@page { size: " + m_Constant.PAPER_LETTER_PORTRAIT + " }" + vbCrLf)
            sb.Append(".dotted {border: 2px dotted black; border-style: none none dotted; color: #fff; background-color: #fff; }" + vbCrLf)
            sb.Append("</style>" + vbCrLf)
            sb.Append("</head>" + vbCrLf)
            sb.Append("<body class='" + m_Constant.PAPER_LETTER_PORTRAIT + "'>" + vbCrLf)

            'x.offering_number, x.year, x.person_id, x.korean_name, x.last_name, x.first_name,
            'UPPER(a.street), UPPER(a.city), UPPER(a.province), a.postal_code, a.country, x.total
            Dim total As String = ""

            pbReceipt.Value = 0
            pbReceipt.Visible = True
            pbReceipt.Maximum = rowCount
            pbReceipt.BackColor = Color.LightYellow
            Dim cnt As Integer = 0

            For Each row As DataRow In ds.Tables(0).Rows
                pbReceipt.Value = cnt

                offeringNumber = row("offering_number").ToString
                lastName = row("last_name").ToString
                firstName = row("first_name").ToString
                street = row("street").ToString
                city = row("city").ToString
                postalcode = row("postal_code").ToString
                total = String.Format("{0:C}", row("total"))

                sb.Append("<section class='sheet padding-10mm'>" + vbCrLf)
                sb.Append("<img src='../images/tkpc_receipt_header.png'></img>" + vbCrLf)
                sb.Append("<p>" + vbCrLf)
                sb.Append("<table>" + vbCrLf)
                sb.Append("<tr>" + vbCrLf)
                sb.Append("	 <td width='5%'>Dear</td>" + vbCrLf)
                sb.Append("  <td><div style='font-weight:bold; font-size:17px'>" + firstName + " " + lastName + "</div></td>" + vbCrLf)
                sb.Append("</tr>" + vbCrLf)
                sb.Append("<tr><td colspan='2'>&nbsp;</td></tr>" + vbCrLf)
                sb.Append("<tr><td colspan='2'>We thank you so much for your donations that enable the ministry to continue and grow. We gratefully acknowledge your financial support for the following funds for the year ending December 31, " + selectedYear + ".</td></tr>" + vbCrLf)
                sb.Append("<tr><td colspan='2'>&nbsp;</td></tr>" + vbCrLf)
                sb.Append("<tr>" + vbCrLf)
                sb.Append("  <td width='15%'><div style='font-weight:bold; font-size:17px'>Offering No: </td>" + vbCrLf)
                sb.Append("  <td>" + offeringNumber + "</td>" + vbCrLf)
                sb.Append("</tr>" + vbCrLf)
                sb.Append("<tr>" + vbCrLf)
                sb.Append("  <td width='15%'><div style='font-weight:bold; font-size:17px'>Amount: </td>" + vbCrLf)
                sb.Append("  <td>$" + total + "</td>" + vbCrLf)
                sb.Append("</tr>" + vbCrLf)
                sb.Append("</table>" + vbCrLf)
                sb.Append("<p>" + vbCrLf)

                Dim tableContent As String = getOfferingReceiptInfo(lastName, firstName, street, city, postalCode, offeringNumber, selectedYear, total)
                sb.Append(tableContent)
                sb.Append("<hr class='dotted' />" + vbCrLf)
                sb.Append(tableContent)
                sb.Append("</section>" + vbCrLf)

                cnt += 1
            Next

            If cnt = 0 Then

                total = String.Format("0.0")
                sb.Append("<section class='sheet padding-10mm'>" + vbCrLf)
                sb.Append("<img src='../images/tkpc_receipt_header.png'></img>" + vbCrLf)
                sb.Append("<p>" + vbCrLf)
                sb.Append("<table>" + vbCrLf)
                sb.Append("<tr>" + vbCrLf)
                sb.Append("	 <td width='5%'>Dear</td>" + vbCrLf)
                sb.Append("  <td><div style='font-weight:bold; font-size:17px'>" + firstName + " " + lastName + "</div></td>" + vbCrLf)
                sb.Append("</tr>" + vbCrLf)
                sb.Append("<tr><td colspan='2'>&nbsp;</td></tr>" + vbCrLf)
                sb.Append("<tr><td colspan='2'>We thank you so much for your donations that enable the ministry to continue and grow. We gratefully acknowledge your financial support for the following funds for the year ending December 31, " + selectedYear + ".</td></tr>" + vbCrLf)
                sb.Append("<tr><td colspan='2'>&nbsp;</td></tr>" + vbCrLf)
                sb.Append("<tr>" + vbCrLf)
                sb.Append("  <td width='15%'><div style='font-weight:bold; font-size:17px'>Offering No: </td>" + vbCrLf)
                sb.Append("  <td>" + offeringNumber + "</td>" + vbCrLf)
                sb.Append("</tr>" + vbCrLf)
                sb.Append("<tr>" + vbCrLf)
                sb.Append("  <td width='15%'><div style='font-weight:bold; font-size:17px'>Amount: </td>" + vbCrLf)
                sb.Append("  <td>$" + total + "</td>" + vbCrLf)
                sb.Append("</tr>" + vbCrLf)
                sb.Append("</table>" + vbCrLf)
                sb.Append("<p>" + vbCrLf)

                Dim tableContent As String = getOfferingReceiptInfo(lastName, firstName, street, city, postalCode, offeringNumber, selectedYear, total)
                sb.Append(tableContent)
                sb.Append("<hr class='dotted' />" + vbCrLf)
                sb.Append(tableContent)
                sb.Append("</section>" + vbCrLf)
            End If

            sb.Append("</body>" + vbCrLf)
            sb.Append("</html>" + vbCrLf)
            Dim fileHelper As New FileHelper
            Dim completed As Integer = fileHelper.writeToFile(fileName, sb.ToString)
            If completed = 1 Then
                System.Diagnostics.Process.Start("file:///" & m_Constant.FILE_FOLDER + fileName)
                'wbResult.Navigate(")
            End If

        Catch ex As Exception
            Logger.LogError(ex.ToString)
        Finally
            offeringDAL = Nothing
            ds = Nothing
            sb = Nothing
            Me.Cursor = Cursors.Default
            pbReceipt.Visible = False
        End Try
    End Sub

    Private Function getOfferingReceiptInfo(ByVal lastName As String, ByVal firstName As String, ByVal street As String, ByVal city As String, ByVal postalCode As String, ByVal offeringNumber As String, ByVal selectedYear As String, ByVal amount As String) As String
        Dim sb2 As StringBuilder = New StringBuilder()

        Try
            sb2.Append("<table>" + vbCrLf)
            sb2.Append("<tr>" + vbCrLf)
            sb2.Append("	 <td width='45%'>" + vbCrLf)
            sb2.Append("  <table>" + vbCrLf)
            sb2.Append("		<tr><td><div style='font-weight:bold; font-size:17px'>&nbsp;&nbsp;Your Official Donation Receipt Copy</div></td></tr>" + vbCrLf)
            sb2.Append("		<tr><td><div style='font-weight:bold; font-size:17px'>&nbsp;&nbsp;for Income Tax Purposes</div></td></tr>" + vbCrLf)
            sb2.Append("		<tr><td>&nbsp;</td></tr>" + vbCrLf)
            sb2.Append("		<tr><td>&nbsp;</td></tr>" + vbCrLf)
            sb2.Append("		<tr><td>&nbsp;</td></tr>" + vbCrLf)
            sb2.Append("		<tr><td>&nbsp;</td></tr>" + vbCrLf)
            sb2.Append("		<tr><td><div style='font-weight:bold; font-size:17px'>&nbsp;&nbsp;" + firstName + " " + lastName + "</div></td></tr>" + vbCrLf)
            If String.IsNullOrEmpty(street) = False Then
                sb2.Append("		<tr><td><div>&nbsp;&nbsp;" + street + "</div></td></tr>" + vbCrLf)
            End If

            Dim cityPostalCode As String = ""
            If String.IsNullOrEmpty(city) = False Then
                cityPostalCode = city.Trim
            End If

            If String.IsNullOrEmpty(postalCode) = False Then
                If String.IsNullOrEmpty(cityPostalCode) = False Then
                    cityPostalCode += ", " + postalCode
                End If
            End If

            sb2.Append("		<tr><td><div>&nbsp;&nbsp;" + cityPostalCode + "</div></td></tr>" + vbCrLf)
            sb2.Append("		<tr><td>&nbsp;</td></tr>" + vbCrLf)
            sb2.Append("		<tr><td>&nbsp;</td></tr>" + vbCrLf)
            sb2.Append("		<tr><td>&nbsp;</td></tr>" + vbCrLf)
            sb2.Append("		<tr><td>&nbsp;</td></tr>" + vbCrLf)
            sb2.Append("  </table>" + vbCrLf)
            sb2.Append("  </td>" + vbCrLf)
            sb2.Append("	 <td width='10%'>&nbsp;</td>" + vbCrLf)
            sb2.Append("	 <td width='45%'>" + vbCrLf)
            sb2.Append("	 <table>" + vbCrLf)
            sb2.Append("		<tr><td><div style='font-weight:bold; font-size:17px'>" + txtName.Text + "</div></td></tr>" + vbCrLf)
            sb2.Append("		<tr><td><div>" + txtAddr1.Text + "</div></td></tr>" + vbCrLf)
            sb2.Append("		<tr><td><div>" + txtAddr2.Text + "</div></td></tr>" + vbCrLf)
            sb2.Append("		<tr><td><div>Telephone: " + txtPhone.Text + "</div></td></tr>" + vbCrLf)
            sb2.Append("		<tr><td><div>&nbsp;</div></td></tr>" + vbCrLf)
            sb2.Append("		<tr><td><div>Registration Number: " + txtRegNo.Text + "</div></td></tr>" + vbCrLf)
            sb2.Append("        <tr><td><div>Receipt Number: " + offeringNumber + " - " + selectedYear + "</div></td></tr>" + vbCrLf)
            sb2.Append("		<tr><td><div>Amount: $" + amount + "</div></td></tr>" + vbCrLf)
            sb2.Append("		<tr><td><div>Date of Issue: " + dtpReceiptIssue.Text + "</div></td></tr>" + vbCrLf)
            sb2.Append("		<tr><td><div><img src='" + m_Constant.ROOT_FOLDER + "\images\receipt_signature.png'/></div></td></tr>" + vbCrLf)
            sb2.Append("		<tr><td><div>Treasurer " + txtTreasurer.Text + "</div></td></tr>" + vbCrLf)
            sb2.Append("  </table>" + vbCrLf)
            sb2.Append("  </td>" + vbCrLf)
            sb2.Append("</tr>" + vbCrLf)
            sb2.Append("<tr>" + vbCrLf)
            sb2.Append("<td colspan='3'>For information on all registered charities in Canada under the Income Tax Act, please visit: Canada Revenue Agency www.cra-arc.gc.ca/charities</td>" + vbCrLf)
            sb2.Append("</tr>" + vbCrLf)
            sb2.Append("</table>" + vbCrLf)
        Catch ex As Exception
            Logger.LogError(ex.ToString)
        Finally
        End Try

        getOfferingReceiptInfo = sb2.ToString

    End Function

    Private Sub btCancel_Click(sender As Object, e As EventArgs) Handles btCancel.Click
        Me.Close()
        Me.Dispose()
    End Sub

    Private Sub btOK_Click(sender As Object, e As EventArgs) Handles btOK.Click
        writeOfferingReceipt()

        Dim config As Configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None)

        '<add key = "church_name" value="Toronto Korean Presbyterian Church"/>
        '<add key = "church_addr1" value="67 Scarsdale Rd."/>
        '<add key = "church_addr2" value="North York, ON. M3B 2R2"/>
        '<add key = "church_phone" value="416-477-5963"/>
        '<add key = "church_reg_no" value="82777 7970 RR0001"/>
        '<add key = "treasurer" value="Kim Sung Kyum"/>

        config.AppSettings.Settings(m_Constant.RECEIPT_CHURCH_NAME).Value = txtName.Text
        config.AppSettings.Settings(m_Constant.RECEIPT_CHURCH_ADDR1).Value = txtAddr1.Text
        config.AppSettings.Settings(m_Constant.RECEIPT_CHURCH_ADDR2).Value = txtAddr2.Text
        config.AppSettings.Settings(m_Constant.RECEIPT_CHURCH_PHONE).Value = txtPhone.Text
        config.AppSettings.Settings(m_Constant.RECEIPT_CHURCH_REG_NO).Value = txtRegNo.Text
        config.AppSettings.Settings(m_Constant.RECEIPT_TREASURER).Value = txtTreasurer.Text

        config.Save(ConfigurationSaveMode.Modified)
        ConfigurationManager.RefreshSection("appSettings")

        Me.Close()
        Me.Dispose()
    End Sub
End Class
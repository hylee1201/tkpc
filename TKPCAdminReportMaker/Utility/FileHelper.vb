Imports System.IO
Imports System.Text
Imports Microsoft.Office.Interop.Excel
Imports System.Drawing.Printing
Imports System.Diagnostics
Imports IronPdf
Imports Microsoft.Office.Interop
Imports System.Drawing
Imports System.Data.OleDb

Public Class FileHelper
    Public Function generateExcel(ds As DataSet) As String
        Dim excel As New Application
        Dim wBook As Workbook = Nothing
        Dim wSheet As Worksheet = Nothing
        generateExcel = ""

        Try
            wBook = excel.Workbooks.Add()
            wSheet = CType(wBook.ActiveSheet(), Worksheet)

            Dim dt As System.Data.DataTable = ds.Tables(0)
            Dim dc As System.Data.DataColumn
            Dim dr As System.Data.DataRow
            Dim colIndex As Integer = 0
            Dim rowIndex As Integer = 0

            For Each dc In dt.Columns
                colIndex = colIndex + 1
                excel.Cells(1, colIndex) = dc.ColumnName
            Next
            For Each dr In dt.Rows
                rowIndex = rowIndex + 1
                colIndex = 0
                For Each dc In dt.Columns
                    colIndex = colIndex + 1
                    excel.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)
                Next
            Next
            wSheet.Columns.AutoFit()
            generateExcel = m_Constant.ROOT_FOLDER + "\offering_by_type.xlsx"
            If System.IO.File.Exists(generateExcel) Then
                System.IO.File.Delete(generateExcel)
            End If

            wBook.SaveAs(generateExcel)
        Catch ex As Exception
            Logger.LogError(ex.ToString)
        Finally
            wBook.Close()
            excel.Quit()
        End Try
        Return generateExcel
    End Function

    Public Function writeToFile(ByVal fileName As String, ByVal content As String) As Integer
        Dim completed As Integer = 0
        Dim path As String = m_Constant.FILE_FOLDER + "\" + fileName
        Dim fs As FileStream = Nothing
        Try
            ' Create or overwrite the file.
            fs = File.Create(path)

            ' Add text to the file.
            Dim info As Byte() = New UTF8Encoding(True).GetBytes(content)
            fs.Write(info, 0, info.Length)
            completed = 1
        Catch ex As Exception
            Logger.LogError(ex.ToString)
        Finally
            fs.Close()
            fs = Nothing
        End Try
        Return completed
    End Function
    Public Sub convertHtmlToPDF(ByVal whichFile As Integer, ByVal html As String, ByVal fileName As String)
        'Dim gc As Pechkin.GlobalConfig = New Pechkin.GlobalConfig
        'With gc
        '    .SetDocumentTitle("Test document")
        '    .SetPaperSize(PaperKind.Letter)
        'End With
        ''Create Converter
        'Dim iPechkin As Pechkin.IPechkin = New SynchronizedPechkin(gc)

        ''Create document configuration object
        'Dim objConfig As Pechkin.ObjectConfig = New Pechkin.ObjectConfig()

        ''And set it up using fluent notation too
        'With objConfig
        '    .SetAllowLocalContent(True)
        '    .SetPageUri("@" + html)
        'End With
        'Dim pdfContent As Byte() = iPechkin.Convert(objConfig)

        '----OpenHtmlToPdf 
        'Dim pdfContent As Byte() = Nothing
        'If whichFile = m_Constant.FILE_SURVEY Then
        '    pdfContent = OpenHtmlToPdf.Pdf.From(html).WithGlobalSetting(m_Constant.PAPER_SIZE_LETTER, m_Constant.PAPER_SIZE_LETTER).WithGlobalSetting(m_Constant.PAPER_ORIENTATION_LANDSCAPE, m_Constant.PAPER_ORIENTATION_LANDSCAPE).Content
        'ElseIf whichFile = m_Constant.FILE_MEMBERSHIP Then
        '    pdfContent = OpenHtmlToPdf.Pdf.From(html).WithGlobalSetting(m_Constant.PAPER_SIZE_LETTER, m_Constant.PAPER_SIZE_LETTER).WithGlobalSetting(m_Constant.PAPER_ORIENTATION_PORTRAIT, m_Constant.PAPER_ORIENTATION_PORTRAIT).Content
        'End If

        'Create a PDF from an existing HTML using C#
        Try
            Dim renderer As IronPdf.HtmlToPdf = New IronPdf.HtmlToPdf()
            Dim PDF = renderer.RenderHTMLFileAsPdf(fileName)
            Dim pdfFileName As String = fileName.Replace(m_Constant.FILE_EXTENSION_HTML, m_Constant.FILE_EXTENSION_PDF)

            If isFileOpen(pdfFileName) = True Then
                Dim myProcesses() As Process

                'single process variable
                Dim myProcess As Process

                'Get the list of processes
                myProcesses = Process.GetProcesses()

                For Each myProcess In myProcesses
                    Logger.LogInfo(myProcess.ProcessName)
                    If myProcess.ProcessName = m_Constant.PROCESS_ACROBAT Or
                   myProcess.ProcessName = m_Constant.PROCESS_ACROBAT32 Then
                        myProcess.Kill()
                    End If
                Next
            End If

            PDF.SaveAs(pdfFileName)
        Catch ex As Exception
            MessageBox.Show("Error at FileHelper.convertHtmlToPDF : " + ex.Message)
            Logger.LogError(ex.ToString)
        End Try

        'convertHtmlToPDF = writeByteArrayToFile(fileName.Replace(m_Constant.FILE_EXTENSION_HTML, m_Constant.FILE_EXTENSION_PDF), pdfContent)
    End Sub

    Public Function writeByteArrayToFile(ByVal fileName As String, ByVal byteArray As Byte()) As Boolean
        writeByteArrayToFile = False
        'Open File for reading
        Logger.LogInfo(fileName)
        If isFileOpen(fileName) = True Then
            Dim myProcesses() As Process

            'single process variable
            Dim myProcess As Process

            'Get the list of processes
            myProcesses = Process.GetProcesses()

            For Each myProcess In myProcesses
                Logger.LogInfo(myProcess.ProcessName)
                If myProcess.ProcessName = m_Constant.PROCESS_ACROBAT Or
                   myProcess.ProcessName = m_Constant.PROCESS_ACROBAT32 Then
                    myProcess.Kill()
                End If
            Next
        End If

        Dim fileStream As FileStream = New FileStream(fileName, FileMode.Create, FileAccess.Write)

        Try
            'Writes a block of bytes to this stream using data from  a byte array.
            fileStream.Write(byteArray, 0, byteArray.Length)
            writeByteArrayToFile = True

        Catch ex As Exception
            Logger.LogError(ex.ToString)
        Finally
            'Close File stream
            fileStream.Close()
            fileStream = Nothing
        End Try

        Return writeByteArrayToFile
    End Function

    Public Function deleteFiles(ByVal thresholdDays As Integer, ByVal fileExtension As String) As Integer
        Dim directory As New IO.DirectoryInfo(m_Constant.LOG_FOLDER)
        For Each file As IO.FileInfo In directory.GetFiles
            'MsgBox(file.FullName + ", " + CType((Now - file.CreationTime).Days, String))
            If file.Extension.Equals(fileExtension) AndAlso (Now - file.CreationTime).Days > thresholdDays Then
                Logger.LogInfo(file.FullName + " is being deleted.")
                file.Delete()
            End If
        Next

    End Function

    Public Function isFileOpen(ByVal filename As String) As Boolean
        Dim fOpen As IO.FileStream = Nothing
        Try
            fOpen = System.IO.File.Open(filename, IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.None)
            Return False
        Catch ex As Exception
            Logger.LogError(ex.ToString)
            Return True
        Finally
            If fOpen IsNot Nothing Then
                fOpen.Close()
                fOpen.Dispose()
                fOpen = Nothing
            End If
        End Try
    End Function

    Public Function readNamesFromExcel(ByVal fileName As String) As String
        Dim App As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim Wb As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim wbSheet As Microsoft.Office.Interop.Excel.Worksheet = Nothing
        Dim lastRowIndex As Integer
        Dim currRowIndex As Integer
        Dim cellValue As String
        readNamesFromExcel = ""

        Try
            ' Create a new Excel application instance
            App = New Microsoft.Office.Interop.Excel.Application()
            ' Open a workboob (i.e. Excel file)
            Wb = App.Workbooks.Open(fileName)
            wbSheet = CType(Wb.Sheets.Item(1), Worksheet)

            ' Here's a trick to find last row in this sheet
            lastRowIndex = wbSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell).Row
            currRowIndex = 1

            mdiTKPC.tspbTKPC.Value = 0
            mdiTKPC.tspbTKPC.Visible = True
            mdiTKPC.tspbTKPC.Maximum = lastRowIndex - 1
            mdiTKPC.tspbTKPC.BackColor = Color.LightCoral

            While (currRowIndex < lastRowIndex)
                mdiTKPC.tspbTKPC.Value = currRowIndex
                ' Get one row at time. Now Cells' row index is alway 1
                'cellValue = CType(currRow.Cells(currRowIndex, 1), Range).Value2.ToString()
                cellValue = wbSheet.Range("A" + CType(currRowIndex, String)).Value.ToString
                readNamesFromExcel += "'" + cellValue + "'"
                If currRowIndex < lastRowIndex - 1 Then
                    readNamesFromExcel += ","
                End If
                currRowIndex += 1
            End While

            Logger.LogInfo(readNamesFromExcel)

        Catch ex As Exception
            MessageBox.Show("Error at FileHelper.readExcel " + ex.Message)
            Logger.LogError(ex.ToString)
            Logger.LogInfo(readNamesFromExcel)
            MessageBox.Show(readNamesFromExcel)
        Finally
            If App IsNot Nothing Then
                App.DisplayAlerts = False
            End If

            If Wb IsNot Nothing Then
                Wb.Save()
                Wb.Close()
            End If

            If App IsNot Nothing Then
                App.Quit()
            End If

            ' Dispose all objects
            releaseObject(App)
            releaseObject(Wb)
            releaseObject(wbSheet)
        End Try

        Return readNamesFromExcel
    End Function

    Public Function readFromCSV(ByVal flag As Integer, ByVal fileName As String) As String
        If String.IsNullOrEmpty(fileName) = True Then
            Return "-1"
        End If

        readFromCSV = ""
        Dim TextFileReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(fileName)
        Dim lineCount As Integer = IO.File.ReadAllLines(fileName).Length

        Try
            TextFileReader.TextFieldType = FileIO.FieldType.Delimited
            TextFileReader.SetDelimiters(",")

            Dim TextFileTable As DataTable = Nothing
            Dim CurrentRow As String()

            If TextFileReader.EndOfData Then
                MsgBox("데이터가 없는 빈 화일입니다.", Nothing, "주의")
                TextFileReader.Close()
                TextFileReader.Dispose()
                releaseObject(TextFileReader)
                Exit Function
            End If

            mdiTKPC.tspbTKPC.Value = 0
            mdiTKPC.tspbTKPC.Visible = True
            mdiTKPC.tspbTKPC.Maximum = lineCount
            mdiTKPC.tspbTKPC.BackColor = Color.LightCoral

            While Not TextFileReader.EndOfData
                mdiTKPC.tspbTKPC.Value += 1
                Try
                    CurrentRow = TextFileReader.ReadFields()
                    If Not CurrentRow Is Nothing Then
                        readFromCSV += "'" + CurrentRow(0).ToString() + "',"
                    End If
                Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                    MsgBox("Line " & ex.Message & "에 데이터 형태가 올바르지 않습니다. 건너뛰고 다음 데이터를 읽습니다.", Nothing, "주의")
                End Try
            End While

            If readFromCSV.EndsWith(",") = True Then
                readFromCSV = readFromCSV.Substring(0, readFromCSV.Length - 1)
            End If

            Logger.LogInfo(readFromCSV)
            'MessageBox.Show(readFromCSV)

        Catch ex1 As IOException
            MessageBox.Show("FileHelper.readFromCSV: 화일을 찾을 수 없습니다. 다시 시도해주십시오.")
            Logger.LogError(ex1.ToString)

        Catch ex As Exception
            MessageBox.Show("FileHelper.readFromCSV: 화일을 읽는데 문제가 발생했습니다. 시스템 관리자에게 문의하시기 바랍니다.")
            Logger.LogError(ex.ToString)
        Finally
            TextFileReader.Close()
            TextFileReader.Dispose()
            releaseObject(TextFileReader)
            mdiTKPC.tspbTKPC.Visible = False
        End Try
        Return readFromCSV
    End Function
End Class

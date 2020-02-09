Imports Microsoft.Office.Interop

Public Class frmAbout
    Private Sub frmResult_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        lblWarning.Text = "본 프로그램은 재능 기부에 의해 만들어진 것입니다" + vbCrLf +
                          "따라서 어떠한 상업적인 목적에 의한 불법적인 복제와 배포를 금합니다. " + vbCrLf +
                          "만일 상업적인 목적에 의해 사용될 경우 " + vbCrLf +
                          "법적인 제재를 받을 수 있을 수 있습니다." + vbCrLf + vbCrLf +
                          "프로그램 수정 문의는 hylee1201g@gmail.com로 해주십시오." + vbCrLf +
                          "프로그램 사용시 에러가 발생했을 경우에는 " + m_Constant.LOG_FOLDER + "에" + vbCrLf +
                          "있는 TKPC_ARM_Log 화일들을 첨부하여 이메일을 보내주시기 바랍니다." + vbCrLf

    End Sub
End Class
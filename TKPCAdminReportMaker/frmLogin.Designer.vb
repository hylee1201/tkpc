<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmLogin
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmLogin))
        Me.txtID = New System.Windows.Forms.TextBox()
        Me.txtPass = New System.Windows.Forms.TextBox()
        Me.lblID = New System.Windows.Forms.Label()
        Me.lblPass = New System.Windows.Forms.Label()
        Me.gbLogin = New System.Windows.Forms.GroupBox()
        Me.btLogin = New System.Windows.Forms.Button()
        Me.btCancel = New System.Windows.Forms.Button()
        Me.pbTKPC = New System.Windows.Forms.PictureBox()
        Me.lklblAccess = New System.Windows.Forms.LinkLabel()
        Me.gbLogin.SuspendLayout()
        CType(Me.pbTKPC, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtID
        '
        Me.txtID.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtID.Location = New System.Drawing.Point(95, 39)
        Me.txtID.Margin = New System.Windows.Forms.Padding(5, 3, 5, 3)
        Me.txtID.Name = "txtID"
        Me.txtID.Size = New System.Drawing.Size(172, 26)
        Me.txtID.TabIndex = 0
        '
        'txtPass
        '
        Me.txtPass.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPass.Location = New System.Drawing.Point(95, 73)
        Me.txtPass.Margin = New System.Windows.Forms.Padding(5, 3, 5, 3)
        Me.txtPass.Name = "txtPass"
        Me.txtPass.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPass.Size = New System.Drawing.Size(172, 26)
        Me.txtPass.TabIndex = 1
        '
        'lblID
        '
        Me.lblID.AutoSize = True
        Me.lblID.Font = New System.Drawing.Font("Malgun Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.lblID.Location = New System.Drawing.Point(39, 42)
        Me.lblID.Margin = New System.Windows.Forms.Padding(5, 0, 5, 0)
        Me.lblID.Name = "lblID"
        Me.lblID.Size = New System.Drawing.Size(50, 17)
        Me.lblID.TabIndex = 2
        Me.lblID.Text = "이메일:"
        '
        'lblPass
        '
        Me.lblPass.AutoSize = True
        Me.lblPass.Font = New System.Drawing.Font("Malgun Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.lblPass.Location = New System.Drawing.Point(53, 77)
        Me.lblPass.Margin = New System.Windows.Forms.Padding(5, 0, 5, 0)
        Me.lblPass.Name = "lblPass"
        Me.lblPass.Size = New System.Drawing.Size(37, 17)
        Me.lblPass.TabIndex = 3
        Me.lblPass.Text = "암호:"
        '
        'gbLogin
        '
        Me.gbLogin.Controls.Add(Me.lklblAccess)
        Me.gbLogin.Controls.Add(Me.btLogin)
        Me.gbLogin.Controls.Add(Me.btCancel)
        Me.gbLogin.Controls.Add(Me.txtPass)
        Me.gbLogin.Controls.Add(Me.lblPass)
        Me.gbLogin.Controls.Add(Me.txtID)
        Me.gbLogin.Controls.Add(Me.lblID)
        Me.gbLogin.Location = New System.Drawing.Point(50, 110)
        Me.gbLogin.Name = "gbLogin"
        Me.gbLogin.Size = New System.Drawing.Size(308, 188)
        Me.gbLogin.TabIndex = 4
        Me.gbLogin.TabStop = False
        '
        'btLogin
        '
        Me.btLogin.Enabled = False
        Me.btLogin.Font = New System.Drawing.Font("Malgun Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.btLogin.Location = New System.Drawing.Point(167, 122)
        Me.btLogin.Name = "btLogin"
        Me.btLogin.Size = New System.Drawing.Size(75, 29)
        Me.btLogin.TabIndex = 3
        Me.btLogin.Text = "로그인"
        Me.btLogin.UseVisualStyleBackColor = True
        '
        'btCancel
        '
        Me.btCancel.Font = New System.Drawing.Font("Malgun Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.btCancel.Location = New System.Drawing.Point(67, 122)
        Me.btCancel.Name = "btCancel"
        Me.btCancel.Size = New System.Drawing.Size(75, 29)
        Me.btCancel.TabIndex = 4
        Me.btCancel.Text = "취소"
        Me.btCancel.UseVisualStyleBackColor = True
        '
        'pbTKPC
        '
        Me.pbTKPC.Image = Global.TKPC.My.Resources.Resources.tkpc
        Me.pbTKPC.Location = New System.Drawing.Point(50, 31)
        Me.pbTKPC.Name = "pbTKPC"
        Me.pbTKPC.Size = New System.Drawing.Size(308, 86)
        Me.pbTKPC.TabIndex = 5
        Me.pbTKPC.TabStop = False
        '
        'lklblAccess
        '
        Me.lklblAccess.AutoSize = True
        Me.lklblAccess.Location = New System.Drawing.Point(173, 158)
        Me.lklblAccess.Name = "lklblAccess"
        Me.lklblAccess.Size = New System.Drawing.Size(66, 18)
        Me.lklblAccess.TabIndex = 5
        Me.lklblAccess.TabStop = True
        Me.lklblAccess.Text = "엑세스 문의"
        '
        'frmLogin
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 18.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(409, 328)
        Me.Controls.Add(Me.pbTKPC)
        Me.Controls.Add(Me.gbLogin)
        Me.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(5, 3, 5, 3)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmLogin"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "토론토 한인장로교회 관리자 프로그램에 오신 걸 환영합니다"
        Me.gbLogin.ResumeLayout(False)
        Me.gbLogin.PerformLayout()
        CType(Me.pbTKPC, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents txtID As System.Windows.Forms.TextBox
    Friend WithEvents txtPass As System.Windows.Forms.TextBox
    Friend WithEvents lblID As System.Windows.Forms.Label
    Friend WithEvents lblPass As System.Windows.Forms.Label
    Friend WithEvents gbLogin As System.Windows.Forms.GroupBox
    Friend WithEvents btLogin As System.Windows.Forms.Button
    Friend WithEvents btCancel As System.Windows.Forms.Button
    Friend WithEvents pbTKPC As PictureBox
    Friend WithEvents lklblAccess As LinkLabel
End Class

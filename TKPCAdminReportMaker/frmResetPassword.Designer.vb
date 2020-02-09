<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmResetPassword
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmResetPassword))
        Me.btSave = New System.Windows.Forms.Button()
        Me.btCancel = New System.Windows.Forms.Button()
        Me.txtOld = New System.Windows.Forms.TextBox()
        Me.lblID = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtNew = New System.Windows.Forms.TextBox()
        Me.lblPass = New System.Windows.Forms.Label()
        Me.txtNew2 = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'btSave
        '
        Me.btSave.Font = New System.Drawing.Font("Malgun Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.btSave.Location = New System.Drawing.Point(182, 152)
        Me.btSave.Name = "btSave"
        Me.btSave.Size = New System.Drawing.Size(75, 29)
        Me.btSave.TabIndex = 6
        Me.btSave.Text = "저장"
        Me.btSave.UseVisualStyleBackColor = True
        '
        'btCancel
        '
        Me.btCancel.Font = New System.Drawing.Font("Malgun Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.btCancel.Location = New System.Drawing.Point(75, 152)
        Me.btCancel.Name = "btCancel"
        Me.btCancel.Size = New System.Drawing.Size(75, 29)
        Me.btCancel.TabIndex = 7
        Me.btCancel.Text = "취소"
        Me.btCancel.UseVisualStyleBackColor = True
        '
        'txtOld
        '
        Me.txtOld.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOld.Location = New System.Drawing.Point(135, 34)
        Me.txtOld.Margin = New System.Windows.Forms.Padding(5, 3, 5, 3)
        Me.txtOld.Name = "txtOld"
        Me.txtOld.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtOld.Size = New System.Drawing.Size(157, 26)
        Me.txtOld.TabIndex = 13
        '
        'lblID
        '
        Me.lblID.AutoSize = True
        Me.lblID.Font = New System.Drawing.Font("Malgun Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.lblID.Location = New System.Drawing.Point(46, 69)
        Me.lblID.Margin = New System.Windows.Forms.Padding(5, 0, 5, 0)
        Me.lblID.Name = "lblID"
        Me.lblID.Size = New System.Drawing.Size(81, 17)
        Me.lblID.TabIndex = 16
        Me.lblID.Text = "새로운 암호:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Malgun Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label1.Location = New System.Drawing.Point(57, 37)
        Me.Label1.Margin = New System.Windows.Forms.Padding(5, 0, 5, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(68, 17)
        Me.Label1.TabIndex = 18
        Me.Label1.Text = "이전 암호:"
        '
        'txtNew
        '
        Me.txtNew.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNew.Location = New System.Drawing.Point(135, 66)
        Me.txtNew.Margin = New System.Windows.Forms.Padding(5, 3, 5, 3)
        Me.txtNew.Name = "txtNew"
        Me.txtNew.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtNew.Size = New System.Drawing.Size(157, 26)
        Me.txtNew.TabIndex = 14
        '
        'lblPass
        '
        Me.lblPass.AutoSize = True
        Me.lblPass.Font = New System.Drawing.Font("Malgun Gothic", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.lblPass.Location = New System.Drawing.Point(40, 101)
        Me.lblPass.Margin = New System.Windows.Forms.Padding(5, 0, 5, 0)
        Me.lblPass.Name = "lblPass"
        Me.lblPass.Size = New System.Drawing.Size(88, 17)
        Me.lblPass.TabIndex = 17
        Me.lblPass.Text = "새로운 암호2:"
        '
        'txtNew2
        '
        Me.txtNew2.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNew2.Location = New System.Drawing.Point(135, 98)
        Me.txtNew2.Margin = New System.Windows.Forms.Padding(5, 3, 5, 3)
        Me.txtNew2.Name = "txtNew2"
        Me.txtNew2.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtNew2.Size = New System.Drawing.Size(157, 26)
        Me.txtNew2.TabIndex = 15
        '
        'frmResetPassword
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(333, 217)
        Me.Controls.Add(Me.txtOld)
        Me.Controls.Add(Me.lblID)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtNew)
        Me.Controls.Add(Me.lblPass)
        Me.Controls.Add(Me.txtNew2)
        Me.Controls.Add(Me.btSave)
        Me.Controls.Add(Me.btCancel)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmResetPassword"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "암호 변경"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btSave As Button
    Friend WithEvents btCancel As Button
    Friend WithEvents txtOld As TextBox
    Friend WithEvents lblID As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents txtNew As TextBox
    Friend WithEvents lblPass As Label
    Friend WithEvents txtNew2 As TextBox
End Class

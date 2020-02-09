<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMaxOfferingNumber
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
        Me.btSave = New System.Windows.Forms.Button()
        Me.txtMax = New System.Windows.Forms.TextBox()
        Me.btCancel = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'btSave
        '
        Me.btSave.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.btSave.Location = New System.Drawing.Point(138, 95)
        Me.btSave.Name = "btSave"
        Me.btSave.Size = New System.Drawing.Size(75, 29)
        Me.btSave.TabIndex = 13
        Me.btSave.Text = "저장"
        Me.btSave.UseVisualStyleBackColor = True
        '
        'txtMax
        '
        Me.txtMax.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMax.Location = New System.Drawing.Point(103, 44)
        Me.txtMax.Margin = New System.Windows.Forms.Padding(5, 3, 5, 3)
        Me.txtMax.Name = "txtMax"
        Me.txtMax.Size = New System.Drawing.Size(90, 26)
        Me.txtMax.TabIndex = 15
        '
        'btCancel
        '
        Me.btCancel.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.btCancel.Location = New System.Drawing.Point(45, 95)
        Me.btCancel.Name = "btCancel"
        Me.btCancel.Size = New System.Drawing.Size(75, 29)
        Me.btCancel.TabIndex = 14
        Me.btCancel.Text = "닫기"
        Me.btCancel.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(129, Byte))
        Me.Label1.Location = New System.Drawing.Point(52, 47)
        Me.Label1.Margin = New System.Windows.Forms.Padding(5, 0, 5, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(50, 17)
        Me.Label1.TabIndex = 16
        Me.Label1.Text = "최대치:"
        '
        'frmMaxOfferingNumber
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 18.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(258, 163)
        Me.Controls.Add(Me.btSave)
        Me.Controls.Add(Me.txtMax)
        Me.Controls.Add(Me.btCancel)
        Me.Controls.Add(Me.Label1)
        Me.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmMaxOfferingNumber"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "헌금봉투번호 최대치 설정"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btSave As Button
    Friend WithEvents txtMax As TextBox
    Friend WithEvents btCancel As Button
    Friend WithEvents Label1 As Label
End Class

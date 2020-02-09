<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmConversion2
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
        Me.tsTop = New System.Windows.Forms.ToolStrip()
        Me.tsbtSave = New System.Windows.Forms.ToolStripButton()
        Me.tsbtGbClose = New System.Windows.Forms.ToolStripButton()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.dtpCurrentStartDate = New System.Windows.Forms.DateTimePicker()
        Me.lblOptionTitle = New System.Windows.Forms.Label()
        Me.tsTop.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'tsTop
        '
        Me.tsTop.BackColor = System.Drawing.SystemColors.ControlLight
        Me.tsTop.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tsTop.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsbtSave, Me.tsbtGbClose, Me.ToolStripSeparator2})
        Me.tsTop.Location = New System.Drawing.Point(0, 0)
        Me.tsTop.Name = "tsTop"
        Me.tsTop.Size = New System.Drawing.Size(284, 26)
        Me.tsTop.TabIndex = 21
        Me.tsTop.Text = "ToolStrip1"
        '
        'tsbtSave
        '
        Me.tsbtSave.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tsbtSave.Image = Global.TKPC.My.Resources.Resources.save
        Me.tsbtSave.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbtSave.Name = "tsbtSave"
        Me.tsbtSave.Size = New System.Drawing.Size(53, 23)
        Me.tsbtSave.Text = "저장"
        '
        'tsbtGbClose
        '
        Me.tsbtGbClose.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.tsbtGbClose.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tsbtGbClose.Image = Global.TKPC.My.Resources.Resources.close1
        Me.tsbtGbClose.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsbtGbClose.Name = "tsbtGbClose"
        Me.tsbtGbClose.Size = New System.Drawing.Size(53, 23)
        Me.tsbtGbClose.Text = "닫기"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(6, 26)
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.dtpCurrentStartDate)
        Me.Panel1.Location = New System.Drawing.Point(21, 71)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(243, 57)
        Me.Panel1.TabIndex = 22
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(21, 18)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(77, 19)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "변경할 날짜:"
        '
        'dtpCurrentStartDate
        '
        Me.dtpCurrentStartDate.Font = New System.Drawing.Font("Calibri", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpCurrentStartDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpCurrentStartDate.Location = New System.Drawing.Point(102, 16)
        Me.dtpCurrentStartDate.Name = "dtpCurrentStartDate"
        Me.dtpCurrentStartDate.Size = New System.Drawing.Size(120, 26)
        Me.dtpCurrentStartDate.TabIndex = 6
        '
        'lblOptionTitle
        '
        Me.lblOptionTitle.AutoSize = True
        Me.lblOptionTitle.Font = New System.Drawing.Font("Calibri", 12.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Underline), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOptionTitle.Location = New System.Drawing.Point(43, 49)
        Me.lblOptionTitle.Name = "lblOptionTitle"
        Me.lblOptionTitle.Size = New System.Drawing.Size(53, 19)
        Me.lblOptionTitle.TabIndex = 23
        Me.lblOptionTitle.Text = "Label1"
        '
        'frmConversion2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 145)
        Me.Controls.Add(Me.lblOptionTitle)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.tsTop)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "frmConversion2"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "데이터 일괄처리 옵션"
        Me.tsTop.ResumeLayout(False)
        Me.tsTop.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents tsTop As ToolStrip
    Friend WithEvents tsbtSave As ToolStripButton
    Friend WithEvents tsbtGbClose As ToolStripButton
    Friend WithEvents ToolStripSeparator2 As ToolStripSeparator
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Label2 As Label
    Friend WithEvents dtpCurrentStartDate As DateTimePicker
    Friend WithEvents lblOptionTitle As Label
End Class

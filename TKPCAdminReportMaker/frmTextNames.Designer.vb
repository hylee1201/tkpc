<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmTextNames
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmTextNames))
        Me.rtxtNames = New System.Windows.Forms.RichTextBox()
        Me.btCopy = New System.Windows.Forms.Button()
        Me.btClose = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'rtxtNames
        '
        Me.rtxtNames.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.rtxtNames.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.rtxtNames.Location = New System.Drawing.Point(14, 11)
        Me.rtxtNames.Name = "rtxtNames"
        Me.rtxtNames.Size = New System.Drawing.Size(381, 178)
        Me.rtxtNames.TabIndex = 0
        Me.rtxtNames.Text = ""
        '
        'btCopy
        '
        Me.btCopy.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btCopy.Location = New System.Drawing.Point(223, 212)
        Me.btCopy.Name = "btCopy"
        Me.btCopy.Size = New System.Drawing.Size(87, 27)
        Me.btCopy.TabIndex = 5
        Me.btCopy.Text = "복사"
        Me.btCopy.UseVisualStyleBackColor = True
        '
        'btClose
        '
        Me.btClose.Font = New System.Drawing.Font("맑은 고딕", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btClose.Location = New System.Drawing.Point(91, 212)
        Me.btClose.Name = "btClose"
        Me.btClose.Size = New System.Drawing.Size(87, 27)
        Me.btClose.TabIndex = 6
        Me.btClose.Text = "닫기"
        Me.btClose.UseVisualStyleBackColor = True
        '
        'frmTextNames
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(409, 260)
        Me.Controls.Add(Me.btCopy)
        Me.Controls.Add(Me.btClose)
        Me.Controls.Add(Me.rtxtNames)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmTextNames"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "이름만 복사"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents rtxtNames As RichTextBox
    Friend WithEvents btCopy As Button
    Friend WithEvents btClose As Button
End Class
